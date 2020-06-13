# 成绩条导出与叠加工作表

成绩条导出与叠加程序，可以自动把表头匹配到每一条学生成绩的前方，以便于分发给学生。

下载地址：<https://github.com/812610357/Transcripts-Distributor/releases/download/v4.3/V4.3.xlsm>

## 截图

图片目前走的是OneDrive，加载略微缓慢。

### 控制面板

![控制面板](https://public.bl.files.1drv.com/y4ml6XTwEmwgaFpOEeumiYWiK_y4vvUIDYerteb2QGOPMayy7UB9zNnS6ITZQbAsDglyc8DsAfF2g6pwHLyctNI8PajP670QozKw7V3tAKY0LfZkjzLPJgFZQcu28MwZMnA5aH19vhuEr-Z_aUMhS1_ZGi5fSNMD5wM9BzSdUl4I96x01pZVWYu3dMvrma_P_WnH09iuuK6RpfghpplOBa34ZYxAfuZ1N53qcARP-gPLw6si5pZ6IywHpnSdL1ZtJ36)

### 原始数据表格

![原始数据](https://public.bl.files.1drv.com/y4mXuHK1rUjDtIJXyT3PW1Y6VfaWrn4f3K_x1tb5JGENO4b1e2jNHyjs5gGc9KcWwLO1k2b7qX9RE3INOUltN7PNuAGXI8IwrwBY1pjDx-HpSUevfzUQ0doOOhHDHEAiLtLHQ15qPE7PrzQeta-h_5GiibsMgs_fNX_ZO3-WvbfK51ytqIsHonly--OsjHieqiQzp5F3_U3pxqG5D6CdHDqIGXgHHDyAmfQjPcM3cpm14Qv16--Ro_fQOofShtYioTf)

### 导出成绩条

![导出](https://public.bl.files.1drv.com/y4myQgh_Xoi3JNLqj4Xp6eBnMMqblmu3MWyMi-n_LXtgewlW4169eVVKz8eme7hbxrSAAnZxvq7oM0Q5lc3E2M7W3pbKWU9narwgDm9CffWgQXRITl7t1n2dHJaVLAPg4ziRXx4hg4-YQL_budj0CJd-L0FGEk4rcuRBYnCl6rgcfXuCYARSXrYJY4V7zXOWPNW1oyNWoYn05kX7ve-uHZRwBNQfTbLEo82BBBX2aDfOgPU7litv5ny6Eia7hs2jmXN)

### 多次成绩叠加

![叠加](https://public.bl.files.1drv.com/y4mxbxX65QsbecrY2CyvKdz4wjPU6xJzAQBY8X9UqYFiXUg-j2OU0QUM3tY9kLhzSnppL-dxGGbnlJwVYY-KAPgvJnFRboCJUeyQpTAG8s0d8-pLQCYuXc4Qmt_aSpkbq3gYbaG_d39D5MNQHN_W61ECfM3DDuNIR0rn4mT6JaFFtZEJaji2YbovYzS-6C0kBU14Ozf97dzaS_ImF0OmRKT3RDj3W1aM6gqhs3t0sB-OOWq58OCUCqEkqXEf3WB2r7H)

## 首次使用方法

1. 首次使用将学生成绩单导入到“导入数据”工作表中。
2. 在“控制面板”工作表中的“表头行数”选项中填入刚刚导入表格的表头行数。
3. 点击绿色按钮，系统会开始检测您导入的数据。请仔细核对“导入行数”、“导入列数”、“数据条数”是否正确，如果不正确，请手动手动输入。
4. 如果制作单次成绩条，建议先点击蓝色按钮（系统可能会出现卡顿），以防止系统出现异常，再点击橙色按钮。
5. 制作完成后系统会自动保存，您也可以把“成绩条”工作表进行备份，方便下次进行叠加。

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

1. 导出的成绩条中，含有表头位置、表头长宽、写入次数等标记（A列、B列、1行），如果需要进行多次叠加，请不要删除这些标记，否则将导致程序异常。
2. 一个工作表最多包含1048576行，最终导出的成绩条工作表行数=(表头行数+写入次数+1)×数据行数，请自行计算所需行数，以免发生溢出。
