很多时候 CATIA 界面出 bug（比如图标乱了、命令点不开），往往是因为“无关紧要”的缓存文件坏了，而你辛苦配置的快捷键和选项其实都在“关键文件”里。

以下是根据你 z:\catia\ENV\CAA\B30\settings 目录下的文件进行的分类：

1. 关键核心文件 (Critical)
这些文件存储了你的个人习惯、快捷键、工具条布局和具体的工程设置。丢失这些等于要重新配置软件。

文件名	重要性	说明
FrameConfig.CATSettings	?????	工具栏布局核心。记录了所有工具条的位置、哪些显示、哪些隐藏。如果这是核心诉求，此文件必保。
FrameGeneral.CATSettings	?????	常规设置核心。存储界面风格（P2/P3）、自定义快捷键（如果在 Customize 中设置的键盘映射）、撤销次数等。
FrameUserAliases.CATSettings	????	命令别名。如果你在命令行用简写（如 "c" 代表 "circle"），都存在这。
Licensing.CATSettings	????	许可证配置。记录该启动哪个 License。
DraftingOptions.CATSettings	????	工程图标准。尺寸、字体、投影角度、图纸标准（ISO/ANSI）等所有 Drafting 选项页的设置。
PartInfrastructure.CATSettings	???	零件设计偏好。如是否启用混合设计、是否自动创建集合体、参考平面大小等。
Sketcher.CATSettings	???	草图偏好。网格设置、几何约束自动创建开关。
Assembly.CATSettings	???	装配偏好。自动更新、约束创建模式。
MeasureSettings.CATSettings	???	测量工具。测量单位、精度、字体大小。
2. 也是设置，但不涉及界面布局 (Domain Options)
这些文件决定了具体功能的默认行为，属于“功能选项”。

STEP.CATSettings / DXF.CATSettings: 数据转换设置（导出 STEP/DWG 的版本和选项）。
Search.CATSettings: 搜索偏好。
Printers.CATSettings: 打印机配置。
VisualizationRepository.CATSettings: 3D 显示精度、背景色（如果你改过背景色，不要删这个）。
Tree.CATSettings: 结构树的显示设置（如果你调过树的字体大小或显示内容）。
3. 无关紧要 / 甚至建议定期清理的文件 (Disposable)
这些文件主要记录“你上次关软件时的状态”。如果删除，下次打开 CATIA 只是窗口会回到默认位置，不会丢失配置。

文件名	建议	说明
DialogPosition.CATSettings	?? 可删	对话框位置。记录各个功能弹窗上次出现在屏幕哪里。如果遇到弹窗跑出屏幕外找不到了，删除这个文件立马解决。
DialogEditStack.CATSettings	?? 可删	输入框历史记录。
CATIAV5Cache.CATSettings	?? 可删	本地缓存管理设置（通常设好了不常动，但没那么关键）。
WarmStart.CATSettings	?? 可删	如果 CATIA 崩溃，再次启动时提示“是否恢复上次会话”，相关信息就在这。
CATAutoLogoff.CATSettings	?? 一般	自动注销设置。
MRU.CATSettings	?? 可删	"Most Recently Used" 最近打开的文件列表。
SearchMRUQueries...	?? 可删	搜索历史记录。