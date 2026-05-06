import { spawnSync } from "node:child_process";
import {
  existsSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
  statSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const wranglerToml = readFileSync(path.join(root, "wrangler.toml"), "utf8");
const workerJs = readFileSync(path.join(root, "worker.js"), "utf8");
const bucketMatch = wranglerToml.match(/bucket_name\s*=\s*["']([^"']+)["']/);
const prefixMatch = workerJs.match(/const\s+key\s*=\s*["']([^"']+)\/["']\s*\+/);

if (!bucketMatch) {
  throw new Error("Could not find bucket_name in wrangler.toml");
}

if (!prefixMatch) {
  throw new Error("Could not find the R2 key prefix in worker.js");
}

const bucket = bucketMatch[1];
const prefix = prefixMatch[1];
const stateKey = `${prefix}/.r2-sync-state.json`;
const npx = process.platform === "win32" ? "npx.cmd" : "npx";

const skippedDirs = new Set([".git", ".github", ".wrangler", "node_modules"]);
const skippedFiles = new Set([
  "worker.js",
  "wrangler.toml",
  "package-lock.json",
  "pnpm-lock.yaml",
  "yarn.lock",
  ".env",
  ".dev.vars",
]);

const repairStatePrefixes = ["8fe54f5", "061f082", "9c94e80", "0b32e2a", "50034c3", "e35cc8a"];
const knownRepairFiles = [
  "010.aru",
  "2026-04-WorkLog.html",
  "2026-05-WorkLog.md",
  "APS&MySQL.pdf",
  "Asprova&Oracle.pdf",
  "AsprovaModelDemo.md",
  "AsprovaStufyNote.html",
  "MySQLBase.html",
  "ServerInstall&Set.html",
  "index.html",
  "index.md",
  "winserver.html",
];
const knownRepairDeletedFiles = ["DailyWorkLog.md", "001.mp4", "001.pptx"];

function shouldSkip(relativePath) {
  const parts = relativePath.split(path.sep);
  if (parts.some((part) => skippedDirs.has(part))) return true;
  if (relativePath === path.join("scripts", "sync-r2.mjs")) return true;

  const base = path.basename(relativePath);
  if (skippedFiles.has(base)) return true;
  if (base.startsWith(".") && base !== ".well-known") return true;

  return false;
}

function collectFiles(dir) {
  const files = [];

  for (const entry of readdirSync(dir)) {
    const absolutePath = path.join(dir, entry);
    const relativePath = path.relative(root, absolutePath);
    if (shouldSkip(relativePath)) continue;

    const stat = statSync(absolutePath);
    if (stat.isDirectory()) {
      files.push(...collectFiles(absolutePath));
    } else if (stat.isFile()) {
      files.push(absolutePath);
    }
  }

  return files;
}

function runGit(args) {
  return spawnSync("git", args, { cwd: root, encoding: "utf8" });
}

function currentCommit() {
  const result = runGit(["rev-parse", "HEAD"]);
  if (result.status !== 0) return null;
  return result.stdout.trim();
}

function isAncestor(commit) {
  const result = runGit(["merge-base", "--is-ancestor", commit, "HEAD"]);
  return result.status === 0;
}

function changedFilesSince(commit) {
  const diff = runGit(["diff", "--name-only", "--diff-filter=ACMRT", `${commit}..HEAD`]);
  if (diff.status !== 0) return null;

  return gitPathsToFiles(diff.stdout);
}

function filesChangedInCurrentCommit() {
  const diff = runGit(["diff-tree", "--no-commit-id", "--name-only", "-r", "--diff-filter=ACMRT", "HEAD"]);
  if (diff.status !== 0) return null;

  return gitPathsToFiles(diff.stdout);
}

function filesChangedInRecentCommits() {
  const revList = runGit(["rev-list", "--max-count=8", "HEAD"]);
  if (revList.status !== 0) return [];

  const commits = revList.stdout
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  if (commits.length < 2) return filesChangedInCurrentCommit() ?? [];

  const oldestCommit = commits.at(-1);
  const diff = runGit(["diff", "--name-only", "--diff-filter=ACMRT", `${oldestCommit}..HEAD`]);
  if (diff.status !== 0) return [];

  return gitPathsToFiles(diff.stdout);
}

function deletedPathsSince(commit) {
  const diff = runGit(["diff", "--name-status", "--diff-filter=DR", `${commit}..HEAD`]);
  if (diff.status !== 0) return [];
  return gitNameStatusToDeletedPaths(diff.stdout);
}

function deletedPathsInCurrentCommit() {
  const diff = runGit(["diff-tree", "--no-commit-id", "--name-status", "-r", "--diff-filter=DR", "HEAD"]);
  if (diff.status !== 0) return [];
  return gitNameStatusToDeletedPaths(diff.stdout);
}

function gitPathsToFiles(stdout) {
  return stdout
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((relativePath) => path.resolve(root, relativePath))
    .filter((absolutePath) => {
      if (!existsSync(absolutePath)) return false;
      if (!statSync(absolutePath).isFile()) return false;
      return !shouldSkip(path.relative(root, absolutePath));
    });
}

function gitNameStatusToDeletedPaths(stdout) {
  return stdout
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .flatMap((line) => {
      const parts = line.split(/\t+/);
      const status = parts[0] || "";
      if (status === "D" && parts[1]) return [parts[1]];
      if (status.startsWith("R") && parts[1]) return [parts[1]];
      return [];
    })
    .filter((relativePath) => !shouldSkip(relativePath));
}

function wrangler(args, options = {}) {
  return spawnSync(npx, ["--yes", "wrangler", ...args], {
    cwd: root,
    stdio: options.stdio ?? "inherit",
    encoding: options.encoding,
  });
}

function loadSyncState() {
  const dir = mkdtempSync(path.join(tmpdir(), "r2-sync-state-"));
  const file = path.join(dir, "state.json");
  const result = wrangler(
    ["r2", "object", "get", `${bucket}/${stateKey}`, "--file", file, "--remote"],
    { stdio: "pipe", encoding: "utf8" },
  );

  if (result.status !== 0 || !existsSync(file)) {
    return null;
  }

  try {
    return JSON.parse(readFileSync(file, "utf8"));
  } catch {
    return null;
  }
}

function uploadFile(file) {
  const relativePath = path.relative(root, file).split(path.sep).join("/");
  const objectKey = `${prefix}/${relativePath}`;
  console.log(`Uploading ${relativePath} -> ${objectKey}`);

  const result = wrangler([
    "r2",
    "object",
    "put",
    `${bucket}/${objectKey}`,
    "--file",
    file,
    "--remote",
  ]);

  if (result.error) {
    throw result.error;
  }

  if (result.status !== 0) {
    throw new Error(`Failed to upload ${relativePath}`);
  }

  verifyObjectUploaded(file, objectKey, relativePath);
}

function objectMatchesFile(file, objectKey) {
  const dir = mkdtempSync(path.join(tmpdir(), "r2-object-check-"));
  const downloadedFile = path.join(dir, "object");
  const result = wrangler(
    ["r2", "object", "get", `${bucket}/${objectKey}`, "--file", downloadedFile, "--remote"],
    { stdio: "pipe", encoding: "utf8" },
  );

  if (result.status !== 0 || !existsSync(downloadedFile)) return false;
  return statSync(downloadedFile).size === statSync(file).size;
}

function verifyObjectUploaded(file, objectKey, relativePath) {
  if (!objectMatchesFile(file, objectKey)) {
    throw new Error(`Uploaded object was not found in R2: ${relativePath}`);
  }
}

function deleteObject(relativePath) {
  const objectKey = `${prefix}/${relativePath.split(path.sep).join("/")}`;
  console.log(`Deleting ${objectKey}`);

  const result = wrangler(
    ["r2", "object", "delete", `${bucket}/${objectKey}`, "--remote"],
    { stdio: "pipe", encoding: "utf8" },
  );

  if (result.error) {
    throw result.error;
  }

  if (result.status !== 0) {
    console.log(`Delete skipped or failed for ${relativePath}: ${result.stderr || result.stdout || "no details"}`);
  }
}

function includeRecentChangedFilesForRepair(files) {
  const byPath = new Map(files.map((file) => [path.relative(root, file), file]));

  for (const file of filesChangedInRecentCommits()) {
    const relativePath = path.relative(root, file);
    if (byPath.has(relativePath)) continue;

    console.log(`Adding recent changed file for repair upload: ${relativePath}`);
    byPath.set(relativePath, file);
  }

  return [...byPath.values()];
}

function includeKnownRepairFiles(files, state) {
  const stateCommit = state?.commit || "";
  if (!repairStatePrefixes.some((prefixValue) => stateCommit.startsWith(prefixValue))) {
    return files;
  }

  const byPath = new Map(files.map((file) => [path.relative(root, file), file]));

  for (const relativePath of knownRepairFiles) {
    if (byPath.has(relativePath)) continue;

    const absolutePath = path.resolve(root, relativePath);
    if (!existsSync(absolutePath) || !statSync(absolutePath).isFile()) continue;
    if (shouldSkip(relativePath)) continue;

    console.log(`Adding known repair upload: ${relativePath}`);
    byPath.set(relativePath, absolutePath);
  }

  return [...byPath.values()];
}

function includeKnownRepairDeletes(deletePaths, state) {
  const stateCommit = state?.commit || "";
  if (!repairStatePrefixes.some((prefixValue) => stateCommit.startsWith(prefixValue))) {
    return deletePaths;
  }

  const byPath = new Set(deletePaths);
  for (const relativePath of knownRepairDeletedFiles) {
    if (shouldSkip(relativePath)) continue;
    console.log(`Adding known repair delete: ${relativePath}`);
    byPath.add(relativePath);
  }

  return [...byPath];
}

function saveSyncState(commit) {
  const dir = mkdtempSync(path.join(tmpdir(), "r2-sync-state-"));
  const file = path.join(dir, "state.json");
  writeFileSync(
    file,
    JSON.stringify(
      {
        commit,
        syncedAt: new Date().toISOString(),
      },
      null,
      2,
    ),
  );

  console.log(`Saving sync state -> ${stateKey}`);
  const result = wrangler([
    "r2",
    "object",
    "put",
    `${bucket}/${stateKey}`,
    "--file",
    file,
    "--remote",
  ]);

  if (result.error) {
    throw result.error;
  }

  if (result.status !== 0) {
    throw new Error("Failed to save R2 sync state");
  }
}

const headCommit = currentCommit();
const state = loadSyncState();
let files;
let deletePaths = [];

if (!headCommit) {
  console.log("Could not read current git commit; syncing all content files.");
  files = collectFiles(root);
} else if (!state?.commit) {
  files = filesChangedInCurrentCommit();
  deletePaths = deletedPathsInCurrentCommit();
  if (files === null) {
    console.log("No previous R2 sync state found and current commit diff is unavailable; syncing all content files.");
    files = collectFiles(root);
  } else {
    console.log(
      `No previous R2 sync state found; current commit ${headCommit.slice(0, 7)} has ${files.length} changed content file(s) to sync.`,
    );
  }
} else if (!isAncestor(state.commit)) {
  files = filesChangedInCurrentCommit();
  deletePaths = deletedPathsInCurrentCommit();
  if (files === null) {
    console.log("Previous R2 sync commit is not in the current history and current commit diff is unavailable; syncing all content files.");
    files = collectFiles(root);
  } else {
    console.log(
      `Previous R2 sync commit is not in the current history; current commit ${headCommit.slice(0, 7)} has ${files.length} changed content file(s) to sync.`,
    );
  }
} else {
  files = changedFilesSince(state.commit);
  deletePaths = deletedPathsSince(state.commit);
  if (files === null) {
    files = filesChangedInCurrentCommit();
    deletePaths = deletedPathsInCurrentCommit();
    if (files === null) {
      console.log("Git diff unavailable; syncing all content files.");
      files = collectFiles(root);
    } else {
      console.log(`Git range diff unavailable; current commit ${headCommit.slice(0, 7)} has ${files.length} changed content file(s) to sync.`);
    }
  } else {
    console.log(
      `Git diff from ${state.commit.slice(0, 7)} to ${headCommit.slice(0, 7)} found ${files.length} changed content file(s) to sync.`,
    );
  }
}

if (files.length === 0) {
  console.log("No changed content files from git diff; repairing by uploading recent changed content files.");
  files = includeRecentChangedFilesForRepair(files);
}

files = includeKnownRepairFiles(files, state);
deletePaths = includeKnownRepairDeletes(deletePaths, state);

if (deletePaths.length === 0) {
  console.log("No content files need deletion.");
} else {
  console.log(`Deleting ${deletePaths.length} object(s) from remote r2://${bucket}/${prefix}/`);
}

for (const relativePath of deletePaths) {
  deleteObject(relativePath);
}

if (files.length === 0) {
  console.log("No content files need upload.");
} else {
  console.log(`Syncing ${files.length} file(s) to remote r2://${bucket}/${prefix}/`);
}

for (const file of files) {
  uploadFile(file);
}

if (headCommit) {
  saveSyncState(headCommit);
}

console.log("R2 sync completed.");
