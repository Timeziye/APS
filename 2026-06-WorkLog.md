---
title: 6月份工作日志
author: TimeYe
tags: ["2026-06-WorkLog"]
---
[TOC]

# 工作日志

|    日期     |                           工作内容                           | 工作时长 |
| :---------: | :----------------------------------------------------------: | :------: |
|    06/01    |                 汉萨APS项目用户测试问题记录                  |    7h    |
|    06/02    |                 汉萨APS项目用户测试问题记录                  |    7h    |
|    06/03    | ①汉萨APS项目用户测试问题记录<br />②验证订单类型为“库存增减量”且订单数量为负数时是否会展开工艺BOM |    7h    |
|    06/04    |                 ①汉萨APS项目用户测试问题记录                 |    7h    |
|    06/05    | ①导入CSP缺料号<br />②会议纪要：和MES开会确定前几天测试发现平台问题的解决方案、未来报工接入方案、未来计划发布接口等内容<br />③测试平台功能：a.工艺BOM导出修复b.制造订单导出修复。这两个上海和常州都没有问题了，而且导出速度比之前快很多，大概15秒左右就可以导出，c.资源表的【基本数量】字段仍然存在<br />④导出上海和常州工艺BOM并汇总到一张表上 |    8h    |
|    06/08    | ①删除模型内除模具以外的旧资源<br />②针对之前补充的CSP工艺BOM，在MES平台进行料号级工艺BOM和订单级工艺BOM补充<br />③标准工艺BOM参数化工时修改<br />④换型时间重新收集<br />⑤资源分派有效条件更新及新增<br />⑥修改切管机外径、长度、数量参数和优先级；资源分派有效条件表达式细化到制造BOM的工序有效条件表达式上<br />⑦CTB-WG，ERP内【折弯数】为0的时候，是没有30工序CTB-WG的 |    7h    |
|    06/09    | ①MES平台导入标准工艺BOM总表<br />②检查一下标准工艺BOM是弯管1-8，焊接1-8的工作单，对应型号表上的弯曲数量是否都有值。——均有值，无空值 ③工艺BOM上在工序CTB-WG设置工序有效条件表达式 ——已添加工序有效条件ME.Item.ItemUser_BendNum>0<br />③OP\|Hose-ZZ 资源量从6改为1，删除相应的生产日历  ——资源量已改为1，生产日历已删除 |    8h    |
|    06/10    | ①排查：提供了CSP工艺BOM但是模型里面显示缺少工艺BOM的情况<br /> |    7h    |
|    06/11    |                      ①汉萨模型功能仿制                       |    7h    |
|    06/12    |    ①Asprova加密狗考试<br />②CSP新增工艺BOM整理及平台导入     |          |
| 06/15-06/18 | ①跟随董老师去常州博瑞学习，主要关于MES,MIS,MOM,APS几个系统的对接 <br />②给董老师安装智能体codex和workbuddy，讲解关于codex和workbuddy的基础使用 |   7h/d   |
|    06/22    | ①标准工艺优化 ：标准工艺包含：弯管或焊接，品目表折弯数是0的品目，跳过30弯管工序、40检验工序、50二次切割工序，即设置工序有效条件ME.Item.ItemUser_BendNum>0<br />②标准工艺维护、重建物料工艺、订单工艺<br />③刘老师的模型拷贝到服务器并更换DBIO链接<br /> |    7h    |
|    06/23    | ①PD\|CTB-59-ZY 平台资源量，调整为2，改为按资源量分派<br />②新增PD\|CTB-59-ZY,资源量2,优先级超过* <br />③修改PD\|CTB-59-ZY(气密性测试机)、OP<br />④标准工艺的资源是：CZPD Hose-21-ZY的改成：CZPD\|Hose-21-ZY<br />⑤需要分析一下这个订单2026022761，在ERP状态是40,部分销货，在APS是20,已生产<br/>确认一下是哪里的原因：<br/>⑥按照用户提供的表格修改资源量，并且资源量>1的改为按资源量分派，指定资源添加备注 |    8h    |
|    06/24    | ①导入常州资源和生产日历<br />②统计3个模型的订单数、工作数<br />③*因目前ERP没有提供品目制造周期，所以没有扣减周期，确实存在差异。 1、品目表加一个LT字段,存储品目[制造周期]，来源为ERP，接口暂不配置，因ERP目前无此项数据 2、模型内订单表PO页签[例外信息2]扣减父品目制造周期。 3、订单表加一个[例外信息3]，在排产完参考，以[例外信息2]的交货期换成制造开始时刻计算。<br /> |    7h    |
|    06/25    | ①排程计算失真原因查找定位，提供解决方案，待配置于模型<br />②平台数据维护<br />③变更标准工艺制造时间为表达式，原制造时间*(工作数量*10%)向上取整: 原制造时间数字部分s* Roundup(me..制造数量* 0.1,0) |    7h    |
|    06/26    | ①修改标准工艺BOM工时并导入平台，重建料号工艺和订单工艺<br />②继06/11汉萨模型功能仿制 |    8h    |
|    06/29    | ①修改工序有效条件配置有误， CTB-H1误配置弯管数>0生效<br />②平台标准工艺BOM中[严格遵守后资源]字段值为数字的统一改为’否‘<br />③本机测试最新模型2.2预排速度<br />④去除除前道CNC工序以外的【后资源】字段的值<br />⑤排查出现虚拟资源的原因：缺少订单工艺BOM<br />⑥统一上海和常州的Hose-ZZ为按资源量分派 |    8h    |
|    06/30    | ①核查工序表和BOM表的工序字段中是否出现了资源<br />②测试不同换型矩阵设置的情况下排产速度 <br />③平台换型矩阵中删除OP\|HOSE-ZZ，工序=HAS-DZ，资源OP\|HOSE-ZZ或CZOP\|HOSE-ZZ，前设置：3m<br />④换型矩阵优化降低排产速度（删后资源限制） |    8h    |

# 06/01

| 序号 | 提出时间 | 问题类型 | 测试问题                                                     | 解决方案                                                     | 责任人     | 状态 | 完成时间 |
| ---- | -------- | -------- | ------------------------------------------------------------ | ------------------------------------------------------------ | ---------- | ---- | -------- |
| 1    | 6月1日   | 数据     | 已撤销的工作单，单号为2023020955的出现在APS中                | ERP接口查找原因并更改代码                                    | 邓         |      |          |
| 2    | 6月1日   | 数据     | csp的92个料号缺工艺BOM                                       | 增补                                                         | 张         |      |          |
| 3    | 6月1日   | 数据     | 现有数据工作数已达10w+，超过当前软件购买的工作数5w。         | 方案一：     ①10状态的工作单，筛选30天内的进行排产（导入视图修改）；     ②工作单单品也要排除掉。     方案二：增补购买软件工作数 | 董、顾     |      |          |
| 4    | 6月1日   | APS配置  | 工作单中存在需求和供应两种类别，需要进行区分。需求类单据只算齐套不排产，供应类单据既算齐套也算排产 | ERP中在工作单行上添加字段【生产类型】，从ERP系统判断传值，用于判断单品或总成。若是总成，作为供应单据，既参与排产，也参与物料齐套计算；若是单品，则作为需求单据只参与物料齐套计算。 | 邓、董     |      |          |
| 5    | 6月1日   | 数据     | APS缺工艺BOM时，物料BOM无法导入                              | 更改导入配置                                                 | 董         |      |          |
| 6    | 6月1日   | 数据     | ERP系统早上9:35进行了同步，但MES系统在凌晨1:00进行同步，二者同步时间不一致，造成数据不一致 | 在正式上线时，确定两个系统同步时间和频率要一致               | 董、周、邓 |      |          |
| 7    | 6月1日   | APS配置  | APS数据供需对应关系存在问题：CZT供给除CHF的所有销售分厂      | 软件分支机构表的CZT【规格1-MRP区域】值需改为1                | 董         | 结束 | 6月2日   |
| 8    | 6月1日   | 平台配置 | 平台资源表的【基本资源量】和【基本数量】字段是否要去重       |                                                              | 周         |      |          |
| 9    | 6月1日   | 平台配置 | 生产日历中的工作日在平台中屏蔽，使不可见                     |                                                              | 周         |      |          |
| 10   | 6月1日   | 平台配置 | 制造订单没有导出功能                                         |                                                              | 周         |      |          |
| 11   | 6月1日   | 平台配置 | 平台生产日历资源量不是1的(如人)要单独设置日历                |                                                              | 张         |      |          |

# 06/02

| 序号 | 提出时间 | 问题类型 | 测试问题                                                     | 解决方案                                                     | 责任人     | 预计完成时间 | 状态          | 完成时间 |
| ---- | -------- | -------- | ------------------------------------------------------------ | ------------------------------------------------------------ | ---------- | ------------ | ------------- | -------- |
| 1    | 6月1日   | 数据     | 已撤销的工作单，单号为2023020955的出现在APS中                | ERP接口查找原因并更改代码                                    | 邓         |              | 关闭          | 6月2日   |
| 2    | 6月1日   | 数据     | csp的92个料号缺工艺BOM                                       | ①增补标准工艺BOM； ②ERP导入型号和标准工艺BOM的对应关系       | 张         | 6月5日       | 进行中        |          |
| 3    | 6月1日   | 数据     | 现有数据工作数已达10w+，超过当前软件购买的工作数5w。         | 方案一： ①10状态的工作单，筛选30天内的进行排产（导入视图修改）； ②工作单单品也要排除掉。 方案二：增补购买软件工作数  经过验证，在控制数据量的同时（销售合同进180天，库存合并，增加分厂备货需求），工作数预计在10万左右。考虑到未来订单量的增长，建议增补到20万的工作数。 | 董、顾     |              | 验证中-方案一 | 6月2日   |
| 4    | 6月1日   | APS配置  | 工作单中存在需求和供应两种类别，需要进行区分。需求类单据只算齐套不排产，供应类单据既算齐套也算排产 | ERP中在工作单行上添加字段【生产类型】，从ERP系统判断传值，用于判断单品或总成。若是总成，作为供应单据，既参与排产，也参与物料齐套计算；若是单品，则作为需求单据只参与物料齐套计算。 | 邓、董     |              | 关闭          | 6月2日   |
| 5    | 6月1日   | 数据     | APS缺工艺BOM时，物料BOM无法导入                              | 更改导入配置                                                 | 董         |              | 关闭          | 6月2日   |
| 6    | 6月1日   | 数据     | ERP系统早上9:35进行了同步，但MES系统在凌晨1:00进行同步，二者同步时间不一致，造成数据不一致 | 在正式上线时，确定两个系统同步时间和频率要一致               | 董、周、邓 | 上线时       | 进行中        |          |
| 7    | 6月1日   | APS配置  | APS数据供需对应关系存在问题：CZT供给除CHF的所有销售分厂      | 软件分支机构表的CZT【规格1-MRP区域】值需改为1                | 董         |              | 关闭          | 6月2日   |
| 8    | 6月1日   | 平台配置 | 平台资源表的【基本资源量】和【基本数量】字段是否要去重       |                                                              | 周         |              | 未开始        |          |
| 9    | 6月1日   | 平台配置 | 生产日历中的工作日在平台中屏蔽，使不可见                     |                                                              | 周         |              | 未开始        |          |
| 10   | 6月1日   | 平台配置 | 制造订单没有导出功能                                         |                                                              | 周         |              | 未开始        |          |
| 11   | 6月2日   | 平台配置 | 平台生产日历资源量不是1的(如人)要单独设置日历                |                                                              | 张         | 上线时       |               |          |
| 12   | 6月2日   | APS配置  | 需求数可按30天/60天/90天查看                                 |                                                              | 董         |              |               |          |
| 13   | 6月2日   | 数据     | 分厂手工采购单作为需求纳入齐套分析（以PSR开头；三个分厂+常州+HPU） | IT先在原有采购单上新增数据并做标识，然后APS在收到数据后作为需求进行齐套分析 | 邓、董     |              |               |          |
| 14   | 6月2日   | 数据     | 销售分厂CZZ、CWH、CRS、CGZ、CBJ、CCD的库存数据无需导入APS    | APS导入库存时进行过滤                                        | 董         |              | 关闭          | 6月2日   |
| 15   | 6月2日   | APS配置  | 齐套计算时程序报错                                           | ①将工作单是单品的单位用量默认为1 ②物料BOM是需求量是0的导入时过滤 | 董         | 6月2日       | 关闭          | 6月2日   |
| 16   | 6月2日   | 数据     | 车间和CHQ总库入库不同步，导致供应单消失，而需求单还存在      | 车间工作单核销和总库成品入库要同步                           | 顾         | 6月2日       |               |          |
| 17   | 6月2日   | APS配置  | 调拨建议时间存在问题                                         | 导出表【345】字段映射改为右交期                              | 董         |              | 关闭          | 6月2日   |
| 18   | 6月2日   | APS配置  | 订单关联要考虑状态，先考虑生产中齐套的，再考虑生产中的不齐套 | 部分发货与已生产设定为同一优先级，10优先级最低——品目输入关联排序表达式 | 董         | 6月2日       | 关闭          | 6月2日   |
| 19   | 6月2日   | 数据     | 采购单的单位要从接口转换，从毫米转换到米                     | 检查ERP系统                                                  | 邓         |              | 进行中        |          |

# 06/03

添加分支机构表【规格1】

![image-20260603093811167](https://img.tynote.cn/img/typora/20260603093818244.png#800w)

订单表新建对象属性表格列

![image-20260603094314073](https://img.tynote.cn/img/typora/20260603094314142.png#800w)

# 06/04

![image-20260604170215512](https://img.tynote.cn/img/typora/20260604170215591.png)

## 汉萨项目

### 初始化DBIO链接

类定义表->右击项目->新属性定义，新建字符串数组用于存储不同的DBIO链接值

![image-20260608212301497](https://img.tynote.cn/img/typora/20260608212301588.png#800w)

项目设置->自定义页签->右击[DBIO链接]->属性定义->添加数组子名，为各个数组元素命名

![image-20260608212451676](https://img.tynote.cn/img/typora/20260608212451764.png#800w)

将导入/导出表中对应的连接字符串的值赋值给相应的数组元素

![image-20260608213227544](https://img.tynote.cn/img/typora/20260608213227638.png#800w)

计划参数表->右击[DBIO链接]->控制，进行循环计数器的配置，计数器初始值初始化为1，计数器最大值初始化为PropCount(PROJECT.Child['DBIO'].Child)，即导入/导出表的DBIO表格的数量，循环条件式默认，即循环条件为计数器当前值<=计数器最大值，循环处理表达式默认，即每次递增1。

![image-20260608215422621](https://img.tynote.cn/img/typora/20260608215422713.png#800ws)

所以[初始化DBIO链接]的计数器存储相当于一个数组，最大值是导入/导出表的个数。

计划参数表->右击[DBIO链接]-属性编辑->在通用表达式中写筛选和赋值表达式，若当前遍历的DBIO表格的字段【来源】的值是ERP，就赋值ERP数据库的连接字符串，MES、APS同理。

```
PROJECT.Child['DBIO'].Child表示整个导入/导出表，相当于数组a，
HOLDER.Parent.Command_LoopCounter[1]表示当前计数器的值，   
PROJECT.Child['DBIO'].Child[HOLDER.Parent.Command_LoopCounter[1]]，相当于数组a_[x]，
若当前计数器值为20，则等同于数组a_[20]，也就是遍历到导入/导出表的第20号表，
.DBIOUser_Origin='ERP'表示如果当前表的【来源】字段的值是'ERP'，
.ConnectString=PROJECT.ProjectUser_DBIOLink[1]则赋值它的【连接字符串】的值为数组元素[DBIO链接]_[1]的值
```

```c++
If(
PROJECT.Child['DBIO'].Child[HOLDER.Parent.Command_LoopCounter[1]].DBIOUser_Origin='ERP',PROJECT.Child['DBIO'].Child[HOLDER.Parent.Command_LoopCounter[1]].ConnectString=PROJECT.ProjectUser_DBIOLink[1],
PROJECT.Child['DBIO'].Child[HOLDER.Parent.Command_LoopCounter[1]].DBIOUser_Origin='MES',PROJECT.Child['DBIO'].Child[HOLDER.Parent.Command_LoopCounter[1]].ConnectString=PROJECT.ProjectUser_DBIOLink[2],
PROJECT.Child['DBIO'].Child[HOLDER.Parent.Command_LoopCounter[1]].DBIOUser_Origin='APS',PROJECT.Child['DBIO'].Child[HOLDER.Parent.Command_LoopCounter[1]].ConnectString=PROJECT.ProjectUser_DBIOLink[3],
FALSE
)
#解释
IF(
PROJECT.子对象['DBIO'].子对象[HOLDER.父对象.Command_LoopCounter[1]].DBIOUser_Origin='ERP',PROJECT.子对象['DBIO'].子对象[HOLDER.父对象.Command_LoopCounter[1]].ConnectString=PROJECT.'[DBIO链接]'[1],
PROJECT.子对象['DBIO'].子对象[HOLDER.父对象.Command_LoopCounter[1]].DBIOUser_Origin='MES',PROJECT.子对象['DBIO'].子对象[HOLDER.父对象.Command_LoopCounter[1]].ConnectString=PROJECT.'[DBIO链接]'[2],
PROJECT.子对象['DBIO'].子对象[HOLDER.父对象.Command_LoopCounter[1]].DBIOUser_Origin='APS',PROJECT.子对象['DBIO'].子对象[HOLDER.父对象.Command_LoopCounter[1]].ConnectString=PROJECT.'[DBIO链接]'[3],
FALSE
)
```



### 供需关联

汉萨项目要求只在本工厂内部进行供需关联，跨工厂提供调拨建议

供需关联：供应单和需求单关联。单品是需求单，总成是供应单，排程排的是供应单，因为供需关系只用作物料齐套计算。

标准的产销业务流程是①客户下单(ERP录入销售订单)->②MRP运算(销售订单扣库存、拆BOM、计算需制造数量、采购数量等)->下达生产订单(需要自制的部分生成MO工作单，下发车间)

![image-20260609195309646](https://img.tynote.cn/img/typora/20260609195309759.png#800w)

我们将生产订单+库存+采购定义为供应单，将销售订单定义为需求单。要使需求单不参与排程，可将其订单种类设置为`库存(增减量)`，并将其数量设置为负数（订单数量为负数时无法参与排产）。

要通过品目来关联供需订单（汉萨要求只在本工厂内部进行供需关联），所以在品目表-关联条件式中填写如下表达式

![image-20260604143751107](https://img.tynote.cn/img/typora/20260604143758230.png#800w )



<font color='red'><注>：</font>不要写成`ME.OutputWorkInst_Item.Item_Spec1.SpecUser_Spec1_MrpArea==OTHER.InputWorkInst_Item.Item_Spec1.SpecUser_Spec1_MrpArea`，即`ME.品目.分支机构.'[规格1-MRP区域]'==OTHER.品目.分支机构.'[规格1-MRP区域]'`，前者是按照品目.分支机构.'[规格1-MRP区域]将该工作输入指令的输入品目与其他工作输出指令的输出品目关联起来，而`ME.Order.Spec1.Spec1_MrpArea==OTHER.Order.Spec1.Spec1_MrpArea`，即`ME.订单.分支机构.'[规格1-MRP区域]'==OTHER.订单.分支机构.'[规格1-MRP区域]'`是按照品目.分支机构.'[规格1-MRP区域]将该工作所属的订单与其他工作所属的订单关联起来。也就是一个是按照条件关联工序间的品目，一个是按照条件关联订单。



### 齐套计算

```sql
USE sim_hansaflex;
GO	
DROP TABLE t_MOrder;
CREATE TABLE t_MOrder(              -- 单品与总成工作单
[Code] [VARCHAR](100) NULL,          --订单代码
[OrderExType] [VARCHAR](20) NULL DEFAULT '工作单',    --[订单类型]
[ProdAttribute] [VARCHAR](20) NULL,     --生产属性，P为单品，A为总成
[OrderType] [VARCHAR](20) NULL,      --订单种类
[Item] [VARCHAR](100) NULL,          --品目
[OrderQty] [DECIMAL](18, 4) NULL,    --订单数量
[Branch] [VARCHAR](20) NULL,         --分支机构
[LET] [DATETIME] NULL,               --交货期
-- [Crtd] [DATETIME] NULL,              --创建时间
-- [Uptd] [DATETIME] NULL,              --更新时间
-- [Remark] [VARCHAR](1000) NULL        --备注
)

INSERT INTO t_MOrder
(
    [Code],
    [ProdAttribute],
    [OrderType],
    [Item],
    [OrderQty],
    [Branch],
    [LET]
)
VALUES
('14','A','制造','不锈钢压接外螺母',84.0000,'CBJ','2026-07-03 00:00:00'),
('15','A','制造','焊接法兰接头',67.0000,'CCD','2026-07-03 00:00:00'),
('16','A','制造','活动松套法兰',38.0000,'CGZ','2026-07-03 00:00:00'),
('17','A','制造','金属波纹管',5.0000,'CWH','2026-07-10 00:00:00'),
('18','A','制造','特氟龙软管',15.0000,'CZT','2026-09-03 00:00:00'),
('19','A','制造','弯管',35.0000,'HPU','2026-07-07 00:00:00'),
('20','A','制造','金属波纹管',20.0000,'CBJ','2026-07-10 00:00:00'),
('21','A','制造','特氟龙软管',76.0000,'CCD','2026-09-03 00:00:00'),
('22','A','制造','弯管',123.0000,'CGZ','2026-07-07 00:00:00'),
('1','P','库存(增减量)','金属波纹管',-5.0000,'CBJ','2026-07-10 00:00:00'),
('2','P','库存(增减量)','弯管',-10.0000,'CCD','2026-07-07 00:00:00'),
('3','P','库存(增减量)','特氟龙软管',-10.0000,'CGZ','2026-09-03 00:00:00'),
('4','P','库存(增减量)','金属波纹管',-13.0000,'CHF','2026-07-10 00:00:00'),
('5','P','库存(增减量)','弯管',-21.0000,'CHQ','2026-07-07 00:00:00'),
('6','P','库存(增减量)','特氟龙软管',-34.0000,'CQD','2026-09-03 00:00:00'),
('7','P','库存(增减量)','金属波纹管',-57.0000,'CRS','2026-07-10 00:00:00'),
('8','P','库存(增减量)','特氟龙软管',-42.0000,'CSM','2026-09-03 00:00:00'),
('9','P','库存(增减量)','弯管',-219.0000,'CSP','2026-07-07 00:00:00'),
('10','P','库存(增减量)','金属波纹管',-9.0000,'CTB','2026-07-10 00:00:00'),
('11','P','库存(增减量)','不锈钢压接外螺母',-78.0000,'CWH','2026-07-03 00:00:00'),
('12','P','库存(增减量)','焊接法兰接头',-91.0000,'CZT','2026-07-03 00:00:00'),
('13','P','库存(增减量)','活动松套法兰',-38.0000,'HPU','2026-07-03 00:00:00');

CREATE TABLE t_Inv(                  -- 库存单表
[Code] [VARCHAR](100) NULL,          --订单代码
[OrderExType] [VARCHAR](20) NULL DEFAULT '库存单',    --[订单类型]
[ProdAttribute] [VARCHAR](20) NULL,     --生产属性
[OrderType] [VARCHAR](20) NULL,      --订单种类
[Item] [VARCHAR](100) NULL,          --品目
[OrderQty] [DECIMAL](18, 4) NULL,    --订单数量
[Branch] [VARCHAR](20) NULL,         --分支机构
[LET] [DATETIME] NULL,               --交货期
-- [Crtd] [DATETIME] NULL,              --创建时间
-- [Uptd] [DATETIME] NULL,              --更新时间
-- [Remark] [VARCHAR](1000) NULL        --备注
)

INSERT INTO t_Inv
(
    [Code],
    [ProdAttribute],
    [OrderType],
    [Item],
    [OrderQty],
    [Branch],
    [LET]
)
VALUES
('23',NULL,'库存(增减量)','金属波纹管',33.0000,'CHF','2026-06-01 00:00:00'),
('24',NULL,'库存(增减量)','特氟龙软管',17.0000,'CHQ','2026-06-01 00:00:00'),
('25',NULL,'库存(增减量)','弯管',31.0000,'CQD','2026-06-01 00:00:00'),
('26',NULL,'库存(增减量)','金属波纹管',12.0000,'CRS','2026-06-01 00:00:00'),
('27',NULL,'库存(增减量)','特氟龙软管',24.0000,'CSM','2026-06-01 00:00:00'),
('28',NULL,'库存(增减量)','弯管',32.0000,'CSP','2026-06-01 00:00:00'),
('29',NULL,'库存(增减量)','弯管原料2',49.0000,'CTB','2026-06-01 00:00:00');

CREATE TABLE t_POrder(                  -- 采购单表
[Code] [VARCHAR](100) NULL,          --订单代码
[OrderExType] [VARCHAR](20) NULL DEFAULT '库存单',    --[订单类型]
[ProdAttribute] [VARCHAR](20) NULL,     --生产属性
[OrderType] [VARCHAR](20) NULL,      --订单种类
[Item] [VARCHAR](100) NULL,          --品目
[OrderQty] [DECIMAL](18, 4) NULL,    --订单数量
[Branch] [VARCHAR](20) NULL,         --分支机构
[LET] [DATETIME] NULL,               --交货期
-- [Crtd] [DATETIME] NULL,              --创建时间
-- [Uptd] [DATETIME] NULL,              --更新时间
-- [Remark] [VARCHAR](1000) NULL        --备注
)

INSERT INTO t_Porder
(
    [Code],
    [ProdAttribute],
    [OrderType],
    [Item],
    [OrderQty],
    [Branch],
    [LET]
)
VALUES
('30',NULL,'采购','不锈钢压接外螺母',5.0000,'CHF','2026-06-23 00:00:00'),
('31',NULL,'采购','特氟龙软管原料2',25.0000,'CHQ','2026-06-23 00:00:00'),
('32',NULL,'采购','金属波纹管原料2',9.0000,'CQD','2026-06-23 00:00:00');

SELECT * FROM  dbo.t_MOrder;
SELECT * FROM  dbo.t_Inv;
SELECT * FROM  dbo.t_POrder;

CREATE TABLE Kitting_Peg(        -- 用于接收关联表导出的数据。物料供需配对，哪种类型的物料供货单把物料提供给了哪个需求单的哪个工序。
[L_Order] [VARCHAR](100) NULL,     -- 订单(左)
[L_Type] [VARCHAR](10) NULL,       -- 左订单类型
[L_Item] [VARCHAR](100) NULL,      -- 品目
[PegQty] [DECIMAL](18, 4) NULL,    -- 关联数量
[R_Order] [VARCHAR](100) NULL,     -- 订单(右)
[R_Oper] [VARCHAR](100) NULL       -- 右工作
)

SELECT * FROM dbo.Kitting_Peg;
GO

CREATE VIEW [v_OrPegQ] AS -- 用于计算订单关联数量。
SELECT L_Order,SUM(PegQty) Q FROM Kitting_Peg GROUP BY L_Order;
GO 

SELECT * FROM dbo.v_OrPegQ;

CREATE TABLE [dbo].[Kitting_Ins](   --用于接收工作输入指令表。该表记录工作单的每个工序需要的物料及用量
	[ID] [INT] NULL,
	[Oper] [VARCHAR](100) NULL,
	[Item] [VARCHAR](100) NULL,
	[UQ] [DECIMAL](18, 4) NULL,  -- 单位用量。在模型中为（需求数量-已领数量）/订单数量
	[OrCode] [VARCHAR](100) NULL
) 
GO  

CREATE VIEW [dbo].[v_Kitting_PegInv] AS     -- 从[Kitting_Peg]表中统计'库存'类型的可供给数量InvQ
SELECT R_Order,R_Oper AS Oper,L_Item AS Item,SUM(PegQty) InvQ 
FROM Kitting_Peg
WHERE L_Type='库存'
GROUP BY R_Order,R_Oper,L_Item
GO

CREATE VIEW [dbo].[v_Kitting_InsInvQ] AS  -- 库存统计[v_Kitting_PegInv]告知物料需求表[Kitting_Ins]，你工作单为xxx，工序为xxx，物料为xxx的这条数据的ID为xxx库存可供给数量为InvQ
SELECT ID,b.InvQ 
FROM Kitting_Ins a INNER JOIN v_Kitting_PegInv b ON a.Oper=b.Oper AND a.Item=b.Item and a.OrCode=b.R_Order;
GO 

CREATE VIEW [dbo].[v_Kitting_PegER] AS    -- 从[Kitting_Peg]表中统计'采购'+'工作单'类型的可供给数量ERQ
SELECT R_Order,R_Oper AS Oper,L_Item AS Item,SUM(PegQty) ERQ 
FROM Kitting_Peg
WHERE L_Type='采购' OR L_Type='工作单'
GROUP BY R_Order,R_Oper,L_Item
GO

CREATE VIEW [dbo].[v_Kitting_InsERQ] AS  -- 采购+工作单统计[v_Kitting_PegER]告知物料需求表[Kitting_Ins]，你工作单为xxx，工序为xxx，物料为xxx的这条数据的ID为xxx，'采购'+'工作单'类型可供给数量为ERQ
SELECT ID,b.ERQ 
FROM Kitting_Ins a INNER JOIN v_Kitting_PegER b ON a.Oper=b.Oper AND a.Item=b.Item AND a.OrCode=b.R_Order
GO

CREATE VIEW [dbo].[v_Kitting_PegTQ] AS   -- 从[Kitting_Peg]表中统计'库存'+'采购'+'工作单'类型的可供给数量，也就是总供给数量TQ
SELECT R_Order,R_Oper AS Oper,L_Item AS Item,SUM(PegQty) TQ 
FROM Kitting_Peg
WHERE L_Type='采购' OR L_Type='工作单' OR L_Type='库存'
GROUP BY R_Order,R_Oper,L_Item
GO

CREATE VIEW [dbo].[v_Kitting_InsTQ] AS -- [v_Kitting_PegTQ]告知物料需求表[Kitting_Ins]，你工作单为xxx，工序为xxx，物料为xxx的这条数据的ID为xxx，所有类型总的可供给数量为TQ
SELECT ID,b.TQ 
FROM Kitting_Ins a INNER JOIN v_Kitting_PegTQ b ON a.Oper=b.Oper AND a.Item=b.Item AND a.OrCode=b.R_Order
GO

CREATE VIEW [dbo].[v_Kitting_Ins_Or_InvKitQ] AS  -- 库存可供给数量InvQ可满足多少套
SELECT c.ID,ISNULL(d.Q,0) Q FROM Kitting_Ins c LEFT JOIN 
(
SELECT OrCode,Item,SUM(b.InvQ/a.UQ) Q
FROM Kitting_Ins a LEFT JOIN v_Kitting_InsInvQ b ON a.ID=b.ID GROUP BY OrCode,Item
) d ON c.OrCode=d.OrCode AND c.Item=d.Item
GO

CREATE VIEW [dbo].[v_Kitting_Ins_Or_TotalKitQ] AS  -- 总供给数量TQ可满足多少套
SELECT c.ID,ISNULL(d.Q,0) Q FROM Kitting_Ins c LEFT JOIN 
(
SELECT OrCode,Item,SUM(b.TQ/a.UQ) Q
FROM Kitting_Ins a LEFT JOIN v_Kitting_InsTQ b ON a.ID=b.ID GROUP BY OrCode,Item
) d ON c.OrCode=d.OrCode AND c.Item=d.Item
GO
```

【00初始化】是为了清空上一次关联过的订单的关联字段的值。①筛选当前(订单.关联数量)有值的订单②清空所筛选订单的[关联数量]字段的值。

![image-20260629213443380](https://img.tynote.cn/img/typora/20260629213450495.png#800w)

![image-20260629214021147](https://img.tynote.cn/img/typora/20260629214021210.png#800w)

判断齐套的传递顺序是工作输入指令表->工作表->订单表

工作输入指令表增加字段：[关联库存]、[关联预计入]、[关联总供应]、[库存齐套数]、[总供应齐套数]

工作表：

```c++
[投料序]：判断挂在工作或制造任务下的工作输入指令！=In0的是否是投料序，
If(FValid(ME.制造任务.工作输入指令),
MinIF(ME.制造任务.工作输入指令,TARGET.代码!='In0','Y'),
MinIF(ME.工作输入指令,TARGET.代码!='In0','Y'))
    
[完全领料]：在存在有效投料的前提下，所有输入指令的最大需求是否为0，若为0则标记为完全领料（Y），否则为未完全领料（N），而不具备投料的工作直接置为无效（DELETE）
If(ME.'[投料序]'=='Y',If(If(FValid(ME.制造任务.工作输入指令),
MaxIF(ME.制造任务.工作输入指令,TARGET.代码!='In0',TARGET.数量),
MaxIF(ME.工作输入指令,TARGET.代码!='In0',TARGET.数量))==0,'Y','N'),DELETE)
    
[库存齐套数]：计算当前工作在所有有效输入物料中“库存齐套值最小“的那一项，作为该工作的整体库存齐套水平
If(FValid(ME.制造任务.工作输入指令),
MinIF(ME.制造任务.工作输入指令,TARGET.代码!='In0',TARGET.'[库存齐套数]'),
MinIF(ME.工作输入指令,TARGET.代码!='In0',TARGET.'[库存齐套数]'))

[总供应齐套数]：计算当前工作在所有有效输入物料中“总供应齐套值最小”的那一项，作为该工作的整体总共供应齐套水平
If(FValid(ME.制造任务.工作输入指令),
MinIF(ME.制造任务.工作输入指令,TARGET.代码!='In0',TARGET.'[总供应齐套数]'),
MinIF(ME.工作输入指令,TARGET.代码!='In0',TARGET.'[总供应齐套数]'))
```

订单表

```c++
[库存齐套数]：取当前订单下的库存齐套数的最小值作为整个订单的库存齐套值
Min(ME.工作,TARGET.'[库存齐套数]')

[总供应齐套数]:取当前订单下的总供应齐套数的最小值作为整个订单的总供应齐套值
Min(ME.Operations,TARGET.Kitting_TQ)
```

![image-20260701013252281](https://img.tynote.cn/img/typora/20260701013252407.png#800w)

【001导出数据】：导出关联表和工作输出指令

![image-20260701013608421](https://img.tynote.cn/img/typora/20260701013608481.png#800w)

![image-20260701013630416](https://img.tynote.cn/img/typora/20260701013630470.png#800w)

【002导入工作输入指令汇总信息】：导入242:齐套计算-导入-Ins-PegInvQ、243:齐套计算-导入-Ins-PegERQ、244:齐计算-导入-Ins-PegTQ、245:齐套计算-导入-Ins-OrKitInvQ、246:齐套计算-导入-Ins-OrKitTQ

![image-20260701013806549](https://img.tynote.cn/img/typora/20260701013806607.png#800w)



# 06/05

①导入CSP缺料号

②会议纪要：和MES开会确定前几天测试发现平台问题的解决方案、未来报工接入方案、未来计划发布接口等内容

③测试平台功能

测试内容：系统更新了: 1 工艺bom导出失败的问题已经修复, 2 资源表的基本数量字段在编辑时已经移除(数据库字段保留) 3 制造订单新增导出功能.另外标准工艺bom导入的话, 需要把原来的标准工艺BOM删除然后再导入,否则会重复.这个问题后面修改导入机制时会修复.

测试结果：①工艺BOM导出修复③制造订单导出修复。这两个上海和常州都没有问题了，而且导出速度比之前快很多，大概15秒左右就可以导出，②资源表的【基本数量】字段仍然存在

![image-20260606101128827](https://img.tynote.cn/img/typora/20260606101135912.png#800w)

④导出上海和常州工艺BOM并汇总到一张表上

# 06/08

①删除模型内除模具以外的旧资源

将模型内资源与MES平台资源比对，比对结果如下，且删除的资源已在MES平台中进行核实，结论为均为MES中不存在的资源

![image-20260608123846442](https://img.tynote.cn/img/typora/20260608123853586.png#800w)

②针对之前补充的CSP工艺BOM，在MES平台进行料号级工艺BOM补充和订单级工艺BOM生成

③参数化工时修改

APS配置为参数化工时：CTB-BZ，CTB包装工时60s，总数量超过1400件，工时不超过3天；HAS-WG，软管包装工时40S，总数量超过2000件，工时不超过3天；CTB-MW，CTB金属波纹管 包装工时60s，总数量超过1400件，工时不超过3天；——暂时填写在工艺BOM总表的【后设置】字段，待董老师确认后覆盖到【制造】字段

OP|Hose-ZZ 资源量从6改为1，删除相应的生产日历  ——资源量已改为1，生产日历已删除

④换型时间重新收集

Hose04,05,06,07,08,09换型时间改为5min，PD|CTB-43-YB设为5m；——已改到前设置为5m，换型表对应的7条记录已删，待确认

⑤资源分派有效条件更新及新增

⑥修改PD|CTB-42-YB切管机外径、长度、数量参数和优先级；资源分派有效条件表达式细化到制造BOM的工序有效条件表达式上；

——已迁移到制造BOM的工序有效条件表达式上；资源表的分派有效条件表达式暂未删除；待确认



⑦CTB-WG，ERP内【折弯数】为0的时候，是没有30工序CTB-WG的

a. 检查一下标准工艺BOM是弯管1-8，焊接1-8的工作单，对应型号表上的弯曲数量是否都有值。——均有值，无空值

```mysql
SELECT process_bom_category,bend_num FROM t_item where process_bom_category LIKE '%弯管%';
SELECT b.code,b.bend_num FROM t_item b INNER JOIN t_manufacturing_order a ON b.code=a.item AND b.process_bom_category LIKE '%弯管%';
SELECT DISTINCT b.bend_num FROM t_item b INNER JOIN t_manufacturing_order a where  b.code=a.item AND b.process_bom_category LIKE '%弯管%';
SELECT b.code,b.bend_num FROM t_item b INNER JOIN t_manufacturing_order a ON b.code=a.item AND b.process_bom_category LIKE '%焊接%';
SELECT DISTINCT b.bend_num FROM t_item b INNER JOIN t_manufacturing_order a where  b.code=a.item AND b.process_bom_category LIKE '%焊接%';
```

b. 工艺BOM上在工序CTB-WG设置工序有效条件表达式 ——已添加工序有效条件ME.Item.bend_num>0，待确认



