# CATIA Settings Files Description

This document lists the CATIA settings files found in `ENV\CAA\B30\settings` and their purposes in a table format.

## 1. Infrastructure & General (基础架构与通用)
| Filename | Description (English) | Description (Chinese) |
| :--- | :--- | :--- |
| FrameGeneral.CATSettings | General environment settings (UI style, undo limit, help path). | 全局通用设置：界面风格、撤销步数、帮助文档路径 |
| FrameConfig.CATSettings | Workbench and toolbar layout configurations. | 工作台与工具栏布局 |
| FrameUserAliases.CATSettings | User-defined command aliases. | 用户自定义命令别名 |
| DialogPosition.CATSettings | Last positions and sizes of dialog windows. | 对话框上次显示的位置和大小 |
| Licensing.CATSettings | License selection settings. | 许可证勾选状态 |
| Licensing-limited.CATSettings | Similar to Licensing, possibly for specific limited modes. | 受限模式下的许可证设置 |
| Tree.CATSettings | Specification tree display options (orientation, type). | 结构树显示设置：方向、字体等 |
| TreeCustomize.CATSettings | Customization of the tree structure. | 结构树自定义设置 |
| VisualizationRepository.CATSettings | 3D visualization performance (accuracy, anti-aliasing). | 可视化性能：精度、反锯齿、背景色 |
| VisualizationCluster.CATSettings | Visualization settings for clusters/cgr. | 集群或CGR可视化设置 |
| VisuCustomize.CATSettings | Custom visualization settings. | 自定义可视化设置 |
| VisuFilters.CATSettings | Visualization filters settings. | 可视化过滤器设置 |
| Search.CATSettings | Search tool history and preferences. | 搜索工具历史和偏好 |
| SearchFavoriteQueries.CATSettings | Saved favorite search queries. | 收藏的搜索查询 |
| SearchMRUQueries.CATPreferences | Most recently used search queries. | 最近使用的搜索查询记录 |
| MRU.CATSettings | Most Recently Used files list. | 最近打开文件列表 |
| WarmStart.CATSettings | Session recovery data after a crash. | 崩溃后的会话恢复数据 |
| CATAutoLogoff.CATSettings | Automatic session logoff settings. | 自动注销设置 |
| CATMemWarning.CATSettings | Memory usage warning thresholds. | 内存警告阈值设置 |
| SettingsDialog.CATSettings | Settings for the Tools > Options dialog itself. | 选项对话框本身的设置 |
| CATScriptUIWindow.CATPreferences | Macro editor window settings. | 宏编辑器窗口设置 |
| Scripting.CATSettings | Scripting/Macro execution settings. | 脚本和宏执行设置 |
| Printer.CATSettings | Printer configurations. | 打印机配置 |
| Print.CATSettings | Print command settings (banner, layout). | 打印命令设置 |
| Conferencing.CATSettings | Conferencing/Collaboration settings. | 会议/协作设置 |
| SelectSettings.CATSettings | Selection tool preferences (pixel tolerance). | 选择工具偏好：捕捉像素容差 |
| ToolsOptions.CATPreferences | General Tools > Options preferences. | 工具选项通用偏好 |
| URL_SETTING.CATSettings | URL handling settings. | URL处理设置 |
| TabletSupport.CATSettings | Graphics tablet support. | 绘图板支持 |
| VRButtonCustomize.CATSettings | VR device button mapping. | VR设备按钮映射 |
| VRCommands.CATSettings | VR command settings. | VR命令设置 |
| VirtualReality.CATSettings | General VR settings. | VR通用设置 |

## 2. Mechanical Design (机械设计)
| Filename | Description (English) | Description (Chinese) |
| :--- | :--- | :--- |
| PartInfrastructure.CATPreferences | Part Design general settings (Hybrid Design). | 零件设计通用设置：混合设计开关 |
| PartDocument.CATSettings | Settings specific to .CATPart documents. | CATPart文档特定设置 |
| CATPart.CATSettings | Default settings for new Parts. | 新零件默认设置 |
| Sketcher.CATSettings | Sketcher workbench settings (grid, constraints). | 草图工作台设置：网格、约束 |
| Assembly.CATSettings | Assembly Design settings (update, constraints). | 装配设计设置：更新模式、约束 |
| AssemblyDesign.CATPreferences | Additional Assembly preferences. | 额外装配偏好 |
| Constraint.CATSettings | Constraint creation defaults. | 约束创建默认值 |
| DraftingOptions.CATSettings | Drafting standards and general options. | 工程图标准和通用选项 (核心) |
| DraftingOptions.CATPreferences | User preferences for drafting. | 工程图用户偏好 |
| DraftingSession.CATSettings | Current drafting session settings. | 当前工程图会话设置 |
| DraftingSession.CATPreferences | Drafting session preferences. | 工程图会话偏好 |
| DraftingToolbars.CATPreferences | Drafting workbench toolbar layout. | 工程图工具栏布局 |
| CATDrawing.CATSettings | Drawing document specific settings. | 工程图文档特定设置 |
| CATDrwCheckOptionHeader.CATSettings | Drawing checking tool settings. | 工程图检查工具设置 |
| CAT2DLCheckOptionHeader.CATSettings | 2D Layout checking options. | 2D Layout 检查选项 |
| CAT2DLSession.CATPreferences | 2D Layout session preferences. | 2D Layout 会话偏好 |
| CAT2DLToolsOptions.CATSettings | 2D Layout for 3D Design settings. | 3D中的2D布局设置 |
| D2uNewLayout.CATPreferences | Settings for new 2D layouts. | 新2D布局设置 |
| DftNewDrawingEx.CATPreferences | New drawing creation extended preferences. | 新图纸创建扩展偏好 |
| Hole.CATSettings | Hole feature defaults. | 孔特征默认值 |
| FilletMore.CATPreferences | Fillet tool extended preferences. | 倒圆角工具扩展偏好 |
| AutoFilletMore.CATPreferences | Auto-fillet preferences. | 自动倒圆角偏好 |
| MeasureSettings.CATSettings | Measure tool units and accuracy. | 测量工具单位和精度 |
| CkeTolerance.CATSettings | Knowledge advisor tolerances. | 知识工程公差 |
| Knowledge.CATSettings | Knowledge ware settings (parameters, formulas). | 知识工程设置：参数、公式 |
| KnowledgeDialogs.CATPreferences | Dialogs for knowledge ware. | 知识工程对话框设置 |
| ProductStructure.CATSettings | Product Structure workbench settings. | 产品结构工作台设置 |
| ProductToProduct.CATPreferences | Product to Product conversion settings. | 产品间转换设置 |
| SheetMetal.CATSettings | Sheetmetal design defaults (bend radius, K-factor). | 钣金设计默认值：折弯半径、K因子 |
| StructureDesign.CATSettings | Structure Design workbench settings. | 结构设计工作台设置 |
| WeldGeoSettings.CATPreferences | Weld design geometry settings. | 焊接设计几何设置 |
| GBiWFastening.CATSettings | Body in White Fastening settings. | 白车身紧固件设置 |

## 3. Shape & Surface (形状与曲面)
| Filename | Description (English) | Description (Chinese) |
| :--- | :--- | :--- |
| ShapeDesign.CATSettings | Generative Shape Design settings. | 创成式外形设计设置 |
| ShapeDesign.CATPreferences | Preferences for GSD. | GSD偏好 |
| FreeStyleGlobalUserSettings.CATSettings | Freestyle Shaper settings. | 自由曲面设计设置 |
| CATStCmdSurfCurv.CATSettings | Surface curvature analysis settings. | 曲面曲率分析设置 |
| CATStCmdConnectCheckerAnalysis.CATSettings | Connection checker settings. | 连接检查器设置 |
| CATStCmdDisassemble.CATSettings | Disassemble command settings. | 分解命令设置 |
| CATStCrvAnalysis.CATSettings | Curve analysis settings. | 曲线分析设置 |
| CATStNewColorScale.CATSettings | Color scale for analysis. | 分析色标设置 |
| CATDesSettingsCmd.CATPreferences | Imagine & Shape settings. | Imagine & Shape 设置 |
| CATDesCmdSkinModification.CATPreferences | Skin modification preferences. | 蒙皮修改偏好 |
| CATDesContextStateCommand.CATPreferences | Context state command prefs. | 上下文状态命令偏好 |
| CATDesExportObjFilesCmd.CATPreferences | OBJ export for Imagine & Shape. | I&S OBJ导出 |
| CATDesImportObjFilesCmd.CATPreferences | OBJ import settings. | OBJ导入 |
| ICMEditPositionCmd.CATSettings | ICEM Shape Design edit position. | ICEM 形状设计编辑位置 |
| ICMExtIOOptions.CATSettings | ICEM External I/O options. | ICEM 外部I/O选项 |
| DieFaceDesign.CATSettings | Die Face Design settings. | 模具面设计设置 |
| MoldDesignCatalog.CATSettings | Mold Design catalog settings. | 模具设计目录设置 |

## 4. Digital Mockup (DMU) & Simulation
| Filename | Description (English) | Description (Chinese) |
| :--- | :--- | :--- |
| DMUNavigator.CATPreferences | DMU Navigator settings / Cache management. | DMU漫游器/缓存管理 |
| Clash.CATSettings | Clash analysis defaults. | 干涉检查默认值 |
| ClashPublish.CATSettings | Clash result publishing settings. | 干涉结果发布设置 |
| SectioningRepository.CATSettings | Sectioning tool preferences. | 剖面工具偏好 |
| Fitting.CATSettings | DMU Fitting simulation settings. | DMU装配模拟设置 |
| DMUOptimizer.CATSettings | DMU Optimizer settings. | DMU优化器设置 |
| CATIAV5Cache.CATSettings | Cache management for large assemblies. | 大装配缓存管理 |
| DNBSimSimulationSettings.CATSettings | Delmia simulation settings. | Delmia 仿真设置 |
| DNBImportD5.CATSettings | Delmia D5 import settings. | Delmia D5 导入设置 |
| CATSAMAnalysisUI.CATPreferences | Analysis (FEA) UI preferences. | 有限元分析界面偏好 |
| CATElfUserPreferences.CATPreferences | ELFY (FEA) user preferences. | 有限元分析用户偏好 |

## 5. Interoperability & Data Exchange (数据交换)
| Filename | Description (English) | Description (Chinese) |
| :--- | :--- | :--- |
| STEP.CATSettings | STEP export/import options. | STEP 导入导出选项 |
| DXF.CATSettings | DWG/DXF export/import options. | DXF/DWG 导入导出选项 |
| IG2.CATSettings | IGES export/import options. | IGES 导入导出选项 |
| V4Writing.CATSettings | Saving as V4 format options. | 保存为V4格式选项 |
| GeometrybV4ToV5.CATSettings | V4 to V5 geometry migration. | V4到V5几何迁移 |
| SpecifV4ToV5.CATSettings | V4 to V5 specification migration. | V4到V5规范迁移 |
| MigrBatch.CATSettings | Batch migration settings. | 批量迁移设置 |
| MultiCAD.CATSettings | Multi-CAD import settings. | 多CAD导入设置 |
| CAT3DXml.CATSettings | 3DXML format settings. | 3DXML格式设置 |
| CGRFormat.CATSettings | CGR export settings. | CGR格式设置 |
| VrmlFormat.CATSettings | VRML export settings. | VRML格式设置 |
| Report.CATSettings | Report generation settings. | 报告生成设置 |
| CDMAInterop.CATSettings | CDMA interoperability. | CDMA 互操作性 |

## 6. PLM & Collaboration (PLM与协作)
| Filename | Description (English) | Description (Chinese) |
| :--- | :--- | :--- |
| CATEnoviaLCAInterop.CATSettings | Enovia LCA interoperability. | Enovia LCA 互操作性 |
| CATReconcileSettings.CATSettings | Reconciliation settings. | 协调设置 |
| SmarTeamCAI.CATSettings | SmarTeam integration settings (CAI). | SmarTeam 集成 |
| SmarTeamCIX.CATSettings | SmarTeam CIX settings. | SmarTeam CIX 配置 |
| SmarTeamScripts.CATSettings | SmarTeam scripts settings. | SmarTeam 脚本 |
| VPMPSSBlackBox.CATSettings | VPM black box settings. | VPM 黑盒设置 |
| CATIColCollabDesign.CATSettings | Collaborative design settings. | 协同设计 |
| CATIColCollabNetwork.CATSettings | Collaborative network settings. | 协同网络 |
| CATIColConnectivity.CATSettings | Collaborative connectivity. | 协同连接 |
| CATICollabIdentification.CATSettings | Collaboration identification. | 协同身份认证 |
| DocEnv.CATSettings | Document environment settings (DLNames vs Folder). | 文档环境设置 |
| DocView.CATSettings | Document view options. | 文档视图选项 |
| DLNames.CATSettings | Logical path names configuration. | 逻辑路径名配置 |
| DLNamesPreferences.CATPreferences | DLNames user preferences. | DLNames 用户偏好 |
| DLNamesRespositorySearch.CATSettings | DLNames search repo. | DLNames 搜索库 |
| SendTo.CATSettings | "Send To" command (Directory/Mail) settings. | "发送到"命令设置 |
| Publish.CATSettings | Publication settings. | 发布设置 |
| ProjectResourceMngt.CATSettings | Project resource management. | 项目资源管理 |
| SaveMgmtPatternRepos.CATPreferences | Save management pattern/naming rules. | 保存管理命名规则 |

## 7. Catalog & Library (目录与库)
| Filename | Description (English) | Description (Chinese) |
| :--- | :--- | :--- |
| CatalogEditor.CATSettings | Catalog editor settings. | 目录编辑器设置 |
| CatalogBrowser.CATPreferences | Catalog browser preferences. | 目录浏览器偏好 |
| CatalogOption.CATSettings | General catalog options. | 通用目录选项 |
| FdeCreateCatalog.CATSettings | Functional design catalog creation. | 功能设计目录创建 |
| MaterialOptionsSettings.CATSettings | Material library options. | 材质库选项 |
| CATMatPreferencesRepository.CATPreferences | Material preferences. | 材质偏好 |
| LinetypeRepository.CATSettings | Line type definitions. | 线型定义 |
| LightSourceRepository.CATSettings | Light sources. | 光源设置 |
| ThicknessRepository.CATSettings | Thickness definitions. | 厚度定义 |
| DefaultAttributesRepository.CATSettings | Default attributes for unknown types. | 默认属性库 |
| DefaultCreationAttributesPreference.CATPreferences | Creation attributes prefs. | 创建属性偏好 |

## 8. Tolerancing & Annotation (FTA / 3D标注)
| Filename | Description (English) | Description (Chinese) |
| :--- | :--- | :--- |
| CATTPSEditor.CATSettings | Functional Tolerancing & Annotation (FTA) settings. | FTA 设置 |
| CATTPSEditorVisu.CATSettings | FTA visualization settings. | FTA 可视化设置 |
| CATTPSRuleBaseRepository.CATSettings | FTA rule base. | FTA 规则库 |
| CATFTASession.CATPreferences | FTA session preferences. | FTA 会话偏好 |
| CATFcaOptions.CATSettings | Functional analysis options. | 功能分析选项 |

## 9. Specific / Miscellaneous (特定功能与杂项)
| Filename | Description (English) | Description (Chinese) |
| :--- | :--- | :--- |
| 3DCompassAccess.CATPreferences / ... | 3D Compass social/platform access. | 3D罗盘访问设置 |
| 4DNavigator.CATSettings | 4D Navigator settings. | 4D 导航器 |
| AutoFilletMore.CATPreferences | Automated fillet specific prefs. | 自动倒角更多偏好 |
| BColors.CATSettings | Background colors (specific contexts). | 背景色 |
| CColors.CATSettings | Custom colors. | 自定义颜色 |
| CATBehaviorCATFct.CATSettings | Behavior/Function settings. | 行为/功能设置 |
| CATCoCusSettingFormingSettings.CATSettings | Composites forming. | 复材成型 |
| CATCoUINewPlyBookPanel.CATSettings | Composites ply book. | 复材铺层书 |
| CATCompositesSettings.CATSettings | Composites design settings. | 复合材料设计设置 |
| CATFmuPrefRepository.CATPreferences | FMU (Function Mockup Unit) preferences. | FMU 偏好 |
| CATSiMuRepository.CATSettings | Simulation repository. | 仿真库 |
| CATStandardViewSetting.CATSettings | Standard view definitions. | 标准视图定义 |
| CATStatistics.CATSettings | Statistics logging settings. | 统计日志设置 |
| CCDCompatibilitySettings.CATSettings | CADAM compatibility. | CCD/CADAM 兼容性 |
| CT5.CATSettings | Likely CATIA V5 translator/utility. | CT5 相关设置 |
| DiaSettings.CATSettings | Diagram settings. | 图表设置 |
| DressUpDefaultVariantsSettings.CATSettings | Dress-up feature defaults. | 修饰特征默认值 |
| DynLicensing.CATSettings | Dynamic licensing options. | 动态许可证 |
| FRFilletColorization.CATSettings | Fillet recognition colorization. | 倒角识别着色 |
| FoundationDesign.CATSettings | Foundation (ship/plant) design. | 基础设计 |
| GeneralPCS.CATSettings | Performance/Calibration settings. | 性能/校准 |
| GeometricModeler.CATSettings | Kernel modeler settings. | 几何建模核心设置 |
| LayersFilter.CATSettings | Layer filter definitions. | 图层过滤器 |
| LowGraph.CATSettings | Low level graphics settings. | 底层图形设置 |
| NodeCustomize.CATSettings | Node customization. | 节点自定义 |
| NumberDisplay.CATSettings | Number display format (decimals). | 数值显示格式 |
| PartComparison.CATPreferences | Part comparison tool. | 零件比对工具 |
| PartSupply.CATSettings | Part Supply integration. | 零部件供应集成 |
| PenetrationManagement.CATPreferences | Penetration management (piping). | 贯穿管理 |
| SymbolicLinks.CATSettings | Symbolic link handling. | 符号链接处理 |
| TubingMigr.CATSettings | Tubing migration. | 管路迁移 |
| catalog/part/cke.CATSettings | Often backups or aliases for main settings. | 通常是主设置的备份或别名 |
