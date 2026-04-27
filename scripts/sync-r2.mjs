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

  return diff.stdout
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

if (!headCommit) {
  console.log("Could not read current git commit; syncing all content files.");
  files = collectFiles(root);
} else if (!state?.commit) {
  console.log("No previous R2 sync state found; syncing all content files.");
  files = collectFiles(root);
} else if (!isAncestor(state.commit)) {
  console.log("Previous R2 sync commit is not in the current history; syncing all content files.");
  files = collectFiles(root);
} else {
  files = changedFilesSince(state.commit);
  if (files === null) {
    console.log("Git diff unavailable; syncing all content files.");
    files = collectFiles(root);
  } else {
    console.log(
      `Git diff from ${state.commit.slice(0, 7)} to ${headCommit.slice(0, 7)} found ${files.length} changed content file(s) to sync.`,
    );
  }
}

if (files.length === 0) {
  console.log("No changed content files to sync. R2 content upload skipped.");
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
