# 成绩条导出与叠加工作表

成绩条导出与叠加程序，可以自动把表头匹配到每一条学生成绩的前方，以便于分发给学生。

下载地址：<https://github.com/812610357/Transcripts-Distributor/releases/download/v4.1/V4.1.xlsm>

## 截图

图片目前走的是OneDrive，加载略微缓慢。

### 控制面板

![控制面板](https://public.bl.files.1drv.com/y4mYUr3oPUBc-sEG_nm49eioKIXQR_CVXRNOYsLwNnq7DCx40tyd_CYeDsaaZybc5zhu-HCFP5W9pZAMAUHEwNXeOVL3ECwoSmh1L1rTpmQCTiJpYtaDpb0BPA-J89-5R7CgWwaMqlvR4QmyEUHf2OmDqPezoY9DPCgSPBbkL7h_sBbgxCuDEjwMgCbngZW3g3TUIe-wY_21MeBd-g79GgiQSFOc78UBy_HfJkeWrZX6ry580JYx1tjFu-wef6oH8Kp)

### 原始数据表格

![原始数据](https://public.bl.files.1drv.com/y4mP3-kwAwcjPwHW7CPzV__GgRGvzqfLgf0z-5Dj5Gi9UVA57Wm2UDEMECKcapLel9pB3fRAq0JGs-w5Wttw44U0BE3MWOAvXfr10ofsd0CP57Vw6r8prQmXEeAvzB-H2oPchWd7fu4zT23q3WvaJ96NOxajHR9YBDEkTdnHr_dk6WJvX91MT1vh-kd4cjz5p8EjbMpVsJtJp39DliGE67ww2wfqt07J1loDJHSW-OgvW6iioz80n3YHfgOcf2W_RZ7)

### 导出成绩条

![导出](https://public.bl.files.1drv.com/y4m1OoVYSpuZcQcZU_PoLgP3XG1iIeRqVT-gDnFs4AUZ_DF37pZUJydJC_hMPwFbvj__HZRXwAIb0iSyHXZSsacrgpZ5OpnDSN7EiVYLvX7-ZVzNABOxM3XM-BOuGb06zOUeXPtf_zxH3JTd3AkPImvMA2nPZItCmfOrVnY32SvnUxCEaimYWoebh98XQ0hlIlb1Zi03mr2PuxFdDFK21WhjljrnkFRN-_uUEAMEA3oltbXnc76JxfyHPeFZPglvrIS)

### 多次成绩叠加

![叠加](https://public.bl.files.1drv.com/y4mUqjigNbrjI-wmVqlGVRlQo5q2yDcTM3WJDu0QQNShVzMD0droenWUqH-lyxMS_EvvVHtza6BFC2Vixf_X9a3Uo_slys4lt-xVCGBk3qOQ-tj6wKoAbfOib9TBR8aR9oaUordgilGyJI61GBmkGxjJYPngevuk8oE-4FclVVVTFZ5iAYQLTa_8_xxAfgdxVqYXK6fa0sU7Nqi6ignFIwb_lS28qfeQm4xflJuSAQBRwRqqyzzZITHpRJ0ickH9eYY)

## 首次使用方法

1. 首次使用将学生成绩单导入到“导入数据”工作表中。
2. 在“控制面板”工作表中的“表头行数”选项中填入刚刚导入表格的表头行数。系统会开始自动检测您导入的表格（系统可能会出现卡顿），请仔细核对“总行数”、“总列数”、“数据行数”（用总行数减去表头行数）是否正确，如果不正确，请手动手动输入。
3. 如果制作单次成绩条，建议先点击蓝色按钮（系统可能会出现卡顿），以防止系统出现异常，再点击橙色按钮。
4. 制作完成后系统会自动保存，您也可以把“成绩条”工作表进行备份，方便下次进行叠加。

## 多次叠加方法（表头必须一致）

1. 如果上一次制作的成绩条还保存在该工作表中，可在导入新数据之后，直接点击橙色按钮即可。
2. 如果您把上一次制作的成绩条额外多出的话，请原封不动把整个工作表原封不动的放入该工作簿，并把该工作表命名为”成绩条“，导入新数据之后，点击橙色按钮即可。

## 温馨提示

1. 请使用Microsoft Office 2013及以上版本的电子表格程序打开此表格。
2. 推荐使用订阅版的Microsoft office 365运行此表格。
3. 任意版本的金山WPS软件或者其他公司出品的办公软件无法运行此程序，请您谅解。
4. 在使用前请打开宏，你可以每次打开此表格后在工具栏下的“安全警告”中点击“启用内容”来打开宏，或者一次性在“文件”=>“选项”=>“信任中心”=>"信任中心设置"=>“宏设置”中把宏设置为启用。

## 异常处理

1. 您再次使用使用该程序导出单次成绩条时，如果因为表头行数与上一次相同，而不修改”表头行数“选项，在导入新数据后，请先点击蓝色按钮。
2. 如果程序异常，请点击蓝色按钮重置系统。

## 额外说明

1. 导出的成绩条中，含有表头位置（位于A列），写入次数（位于B列）等标记，如果需要进行多次叠加，请不要删除这些标记，否则将导致程序异常。
2. 一个工作表最多包含1048576行，最终导出的成绩条工作表行数=(表头行数+写入次数+1)×数据行数，请自行计算所需行数，以免发生溢出。
