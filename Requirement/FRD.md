# FRD

标签（空格分隔）： DND Develop VBA FRD

---

版本
版本|日期|版本描述
---|---|---
V1|2025/06/13|初版，基础文档

## 文档描述
本程序作为习作，要逐步实现将创建DND人物卡片。

## 业务描述
1. 通过在Excel文件(..\CharacterManagementTool.xlsb)中记录人物相关属性,可以区分多个表格.
2. 用户可以在Excel中维护人物相关信息.
3. 可以选择将相关信息导出到Word的模板文件(位于..\Documents\),自动生成Word文件并且填充内容,以提供打印

## 需求细节
### 变量列表
1. 在模块文件Definitions的文件头部定义如下全局变量,方便在整个项目中调用
    - CharacterMasterList as Variant
    - CharacterMemoList as Variant
    - CharacterAttackSpellList as Variant
    - CharacterEquipmentList as Variant
    - Characters as Scripting.Dictionary  

### 类
1. 以表格细节.1中的表格CharacterMasterSchema的结构,创建类CharacterMaster
    需注意CharacterMaster类中除了表格中的字段外,还有三个属性成员:
    - CharacterMemoList As Collection, 成员都是CharacterMemo对象
    - CharacterAttackSpellList As Collection, 成员都是CharacterAttackSpell对象
    - pCharacterEquipmentList As Collection, 成员都是CharacterEquipment对象
2. 以表格细节.2中的表格CharacterMemoSchema的结构,创建类CharacterMemo
3. 以表格细节.3中的表格CharacterAttackSpellSchema的结构,创建类CharacterAttackSpell
4. 以表格细节.4中的表格CharacterEquipmentSchema的结构,创建类CharacterEquipment

### 函数,过程列表
1. Public Function ConvertListObjectToMarkdown(ByVal TableName As String) As String
    通过指定表格名称来获取表格中的内容并且将他们转化为Markdown格式
2. Public Sub Initialize
    作为项目的启动过程, 进行各个全局变量的读取:
    - 在变量列表.1中的4个Variant,从相应的数据表中读取.例如CharacterMasterList,从worksheet shCharacterMaster的列表读取. 语句为CharacterMasterList = shCharacterMaster.ListObject(1).DataBodyRange
    - 初始化字典对象Characters
3. Public Sub Terminate 
    作为项目的结束过程,销毁各个全局变量
    - 在Initialize过程中读取的几个List,通过erase CharacterMasterList销毁
    - 销毁字典对象
4. Private Function GetFieldColumn(ByVal TableName As String, ByVal FieldName As String) As Integer
    辅助函数,根据传入的类和属性名字来查找对应字段在Excel数据表中的列
    规则如下:
	- 表格等于类名,比如我需要查找CharacterMemo中的MemoType属性,就将类名CharacterMemo和属性名MemoType传入这个函数
	- 该函数于是到对应的数据表,比如CharacterMemo就到shCharacterMemo (WorkSheet) 中的ListObject(CharacterMemo)中查看.HeaderRowRange, 然后查找各个字段,当字段名等于MemoType的时候,返回该字段所在列 Range.Column
5. Public Sub ReadCharacters
    在Initialize过程后执行,作为主要过程读取各个数据表,并且生成相应的对象列表
    - 针对CharacterMasterList中的每一条记录创建一个CharacterMaster对象
    - 针对CharacterMemoList中的每一条记录,创建一个CharacterMemo对象,并且添加到相对应的CharacterMaster.CharacterMemoList,关键链接为CharacterMaster.CharacterId == CharacterMemo.CharacterId
    - 针对CharacterAttackSpellList中的每一条记录,创建一个CharacterAttackSpell对象,并且添加到相对应的CharacterMaster.CharacterAttackSpellList, 关键链接为CharacterMaster.CharacterId == Characterattackspell.CharacterId
    - 针对CharacterEquipmentList中的每一条记录,创建一个CharacterEquipment对象,并且添加到相对应的Charactermaster.CharacterEquipmentList,关键链接为CharacterMaster.CharacterId == Characterequipment.CharacterId
    - 针对每一个CharacterMaster对象,将其添加到数据字典Characters中,键为CharacterMaster.CharacterID,值就是CharacterMaster对象本身
    - 需要考虑到数据表中数据为空的情况.如果CharacterMasterList中没有数据则没有创建CharacterMaster对象.如果另外三张从表没有数据,则CharacterMaster相应的Collection中没有记录
6. Public Sub WriteCharacters
    在Terminate过程之前执行,或者在专门的Save命令中调用。用于将当前的Characters对象保存到相应表格中。
    - 首先将关联到的数据表的记录清空,包括：
        - shCharacterMaster.ListObject(1)
        - shCharacterMemo.ListObject(1)
        - shCharacterAttackSpell.ListObject(1)
        - shCharacterEquipment.ListObject(1)
        - shCharacterSpell.ListObject(1)
        - shCharacterSpellSlot.ListObject(1)
    - 清空表格时，采用辅助函数ClearTableRows，能彻底删除所有数据行，避免残留空行。
    - 针对数据字典Characters中的每一个项目,读取相应的CharacterMaster对象,将其中的各个Collection中的记录同样遍历,写入各个数据表中。
        - Collection CharacterMemoList 写入 shCharacterMemo.ListObject(1)
        - Collection CharacterAttackSpellList 写入 shCharacterAttackSpell.ListObject(1)
        - Collection CharacterEquipmentList 写入 shCharacterEquipment.ListObject(1)
        - Collection CharacterSpellList 写入 shCharacterSpell.ListObject(1)
        - Collection CharacterSpellSlots 写入 shCharacterSpellSlot.ListObject(1)
        - CharacterMaster对象本身的其他属性写入shCharacterMaster.ListObject(1)
    - 字段映射与ReadCharacters过程完全对称，确保所有属性都能正确保存和还原。
7. Public Sub CharacterToUI(ByVal CharacterID)
    在ReadCharacters过程执行以后运行该过程,用于将数据字典中的Character的各项数值写到shGeneral页面中的相应字段
    因此该过程运行的前提是已经将各个数据表的内容读取完毕,并且生成了各个Character的数据字典
    传入参数为CharacterID,从数据字典Characters中定位到对应的角色,然后写入shGeneral页面相应的位置:
    - 对于类中除了Collection的各个成员,已经在shGeneral页面中定义了相应的名称, 将各个值写入就可以
        举例: Character.Strength = 6, 则通过shGeneral.Range("Strength") = Character.Strength来完成写入
        假设条件是类中的属性成员名字一定和shGeneral中的名称定义吻合
        需要考虑到两者有差异的情况. 如果发现有类中成员没有在shGeneral中找到对应名称的情况,弹出对话框显示"成员Strength没有找到对应名称,是否继续写入?", 是,否
        如果用户选择是,则忽略当前成员继续完成剩余部分的写入,如果选择否,则终止当前过程.
    - 对于类中的Collection的各个成员, 应用以下逻辑:
        - 如果是CharacterMemoList:
            1. 针对CharacterMemoList的成员, 其每个成员都是具有CharacterID, MemoType和Contents的CharacterMemo对象.
            2. 针对CharacterMemoList的每个成员, 将其MemoType的值作为TargetName,以及Contents作为Content传入过程WriteDataBlockByRange, 并依次将Contents的内容写入MemoType对应的名称所在区域的右侧单元格列中
        - 遍历Character.CharacterAttackSpellList的成员, 
            如果.ItemType = "Attack"则调用WriteAttacksByRange,传入当前的CharacterAttackSpellList成员(类型为CharacterAttackSpell)
            如果.ItemType = "Spell"则调用WriteSpellsByRange,传入当前的CharacterAttackSpellList成员(类型为CharacterAttackSpell)
        - 遍历Character.CharacterEquipmentList的成员,调用WriteEquipmentsByRange,传入当前的CharacterEquipmentList成员(类型为CharacterEquipment)
        - 遍历Character.CharacterSpellList的成员,调用WriteSpellListByRange,传入当前的CharacterSpellList成员(类型为CharacterSpell)
        - 遍历Character.CharacterSpellSlots的成员,调用WriteSpellSlotsByRange,传入当前的CharacterSpellSlots成员(类型为CharacterSpellSlot)
8. Private Function GetPropertiesFromSchema(ByVal SchemaName As String) As Variant
    辅助函数,用于从schema表中获取属性列表
    - 参数SchemaName指定要读取的schema表名,例如"CharacterMasterSchema"
    - 函数会读取指定schema表中"字段"列的所有值
    - 返回一个包含所有字段名的数组
    - 用于CharacterToUI过程中动态获取需要写入UI的属性列表
    - 这样当schema表发生变化时,不需要修改代码就能自动适应新的属性列表
9. Private Function GetMaxCharacterId() As Long
    辅助函数,用于获取当前字典中最大的CharacterId
    - 遍历Characters字典的所有键
    - 返回最大的ID值
    - 如果字典为空,返回0
    - 用于创建新角色时生成新的ID
10. Public Sub PrepareCharacter(ByVal CharacterId As Variant)
    主程序, 根据传入的参数做如下操作:
    - 执行过程Initialize 初始化对象
    - 执行过程ReadCharacters,读取当前文件中所有的数据,并生成相应的对象以及数据字典
    - 判断CharacterID, 如果为数字:
        - 判断该键值是否存在于字典Characters, if Characters.Exists(CharacterID), 如果已经存在,执行过程CharacterToUI(CharacterId);
        - 如果不存在,提示"当前角色不存在"
        - 如果不为数字,表示新增, 调用GetMaxCharacterId获取当前字典中的CharacterID的最大值并且加1, 作为新的Character对象的键值, 在Characters字典中增加一个新的Character对象,然后执行过程CharacterToUI(CharacterId)
    - 执行过程Terminate 关闭对象
11. Public Function ReadBoolean(ByVal Value As Variant) As Boolean
    辅助函数，将Excel单元格中的字符串（如"Y","y","Yes","yes","YES"）转换为Boolean的True，将"N","n","No","NO"或空白转换为Boolean的False。
    示例：MyObj.IsEquipped = ReadBoolean(Cell.Value)
12. Public Function WriteBoolean(ByVal Value As Boolean) As String
    辅助函数，将Boolean值转换为Excel用的字符串。True转换为"Y"，False转换为"N"。
    示例：Cell.Value = WriteBoolean(MyObj.IsEquipped)
13. Public Function GetCharacterIDFromCharacterIDName(ByVal CharacterIDName As String) As Long
    辅助函数，从类似"1 | 王思齐"的文本中提取"|"左边的内容并转为Long类型返回。
    用法示例：GetCharacterIDFromCharacterIDName("1 | 王思齐") 返回 1。
14. Public Sub PrepDataBlockByRange(ByVal TargetName as String)
    该过程用于清理shGeneral页面的数据输出区域.接受参数TargetName一定是一个名称,通过shGeneral.Range(TargetName).CurrentRegion来返回该名称以及其所在数据区域快. 然后:
    - 如果该区域的列数(Range.Columns.Count)大于1, 则:
    - 将该区域从第二列开始的内容清空. 
    举例: 通过TargetName定位到Treasure名称所在的数据区域CurrentRegion为D5:G8, 在执行该过程后,保留D5:D8现有内容,而E5:G8的内容通过.clear的方式清空.
15. Public Sub WriteDataBlockByRange(ByVal TargetName as string, ByVal Content as string)
    该过程用于向shGeneral页面的区域写入内容.接受参数TargetName一定是一个名称,通过shGeneral.Range(TargetName)来返回该名称所在区域. 然后:
    - 通过定位到的区域的.Offset(0,1)定位到其右侧相邻单元格,如果该单元格为空,则写入Content内容;
    - 如果右侧相邻单元格不为空,则向下一格继续判断,为空则写入Content内容,不为空则继续向下执行.依次类推.
16. Public Sub WriteAttacksByRange(ByRef Attack as CharacterAttackSpell)
    辅助过程，将CharacterAttackSpell对象写入shGeneral页面的Attacks区域。
    - 以TargetName = "Attacks"定位shGeneral.Range(TargetName)
    - 从其右侧一列（.Offset(0,1)）开始，向下查找第一个空白单元格
    - 依次写入：
        - .Offset(0,1) = Attack.Name
        - .Offset(0,2) = Attack.AtkBonus
        - .Offset(0,3) = Attack.Damage_Type
        - .Offset(0,4) = WriteBoolean(Attack.Equiped)
        - .Offset(0,5) = WriteBoolean(Attack.Attuned)
17. Public Sub WriteSpellsByRange(ByRef Spell as CharacterAttackSpell)
    辅助过程，将CharacterAttackSpell对象写入shGeneral页面的Spells区域。
    - 以TargetName = "Spells"定位shGeneral.Range(TargetName)
    - 从其右侧一列（.Offset(0,1)）开始，向下查找第一个空白单元格
    - 依次写入：
        - .Offset(0,1) = Spell.Name
        - .Offset(0,2) = Spell.AtkBonus
        - .Offset(0,3) = Spell.Damage_Type
        - .Offset(0,4) = Spell.SpellMemo
18. Public Sub WriteEquipmentsByRange(ByRef Equipment as CharacterEquipment)
    辅助过程，将CharacterEquipment对象写入shGeneral页面的Equipments区域。
    - 以TargetName = "Equipments"定位shGeneral.Range(TargetName)
    - 从其右侧一列（.Offset(0,1)）开始，向下查找第一个空白单元格
    - 依次写入：
        - .Offset(0,1) = Equipment.Name
        - .Offset(0,2) = Equipment.Quantity
        - .Offset(0,3) = WriteBoolean(Equipment.Attuned)
        - .Offset(0,4) = WriteBoolean(Equipment.Equiped)
19. Public Sub WriteSpellListByRange(ByRef Spell as CharacterSpell)
    辅助过程，将CharacterSpell对象写入shGeneral页面的SpellList区域。
    - 以TargetName = "SpellList"定位shGeneral.Range(TargetName)
    - 从其右侧一列（.Offset(0,1)）开始，向下查找第一个空白单元格
    - 依次写入：
        - .Offset(0,1) = Spell.SpellLevel
        - .Offset(0,2) = Spell.Name
        - .Offset(0,3) = Spell.Description
        - .Offset(0,4) = WriteBoolean(Spell.Prepared)
20. Public Sub WriteSpellSlotsByRange(ByRef SpellSlot as CharacterSpellSlot)
    辅助过程，将CharacterSpellSlot对象写入shGeneral页面的SpellSlots区域。
    - 以TargetName = "SpellSlots"定位shGeneral.Range(TargetName)
    - 从其右侧一列（.Offset(0,1)）开始，向下查找第一个空白单元格
    - 依次写入：
        - .Offset(0,1) = SpellSlot.SpellLevel
        - .Offset(0,2) = SpellSlot.SlotsTotal
        - .Offset(0,3) = SpellSlot.SlotsExpended
21. Public Function SpellSlotByLevel(ByVal SpellLv As Integer) As String
    根据SpellLv返回SlotsTotal。
    在CharacterMaster对象的CharacterSpellSlots中遍历记录,如果CharacterSpellSlot对象的SpellLevel=SpellLv,则返回SlotsTotal。因为记录的唯一性,所以只要匹配到SpellLevel=SpellLv,就可以返回SlotsTotal并终止循环。如果没有匹配到SpellLevel或者CharacterSpellSlots为空,则返回空字符串。
22. Public Function SpellSlotsInString(ByVal SpellLv As Integer) As String
    根据SpellLv将法术位以图形方式返回字符串。
    在CharacterMaster对象的CharacterSpellSlots中遍历记录,如果CharacterSpellSlot对象的SpellLevel=SpellLv,则根据SlotsTotal返回字符串,字符串构成为ChrW(&H25EF)重复SlotsTotal次数。例如SlotsTotal=5,则返回◯◯◯◯◯。因为记录的唯一性,所以只要匹配到SpellLevel=SpellLv,就可以返回SlotsTotal并终止循环。如果没有匹配到SpellLevel或者CharacterSpellSlots为空,则返回空字符串。
23. Public Function ExportSpellList(ByVal Pattern As String, ByVal SpellLv As Integer) As String
    根据输入的SpellLv在CharacterSpellList中定位相应的记录(CharacterSpell.SpellLevel = SpellLv,记录不唯一,因此需要遍历整个列表),针对每条匹配到的记录依次取NeedPrep/Prepared, Name, Description, 根据Pattern来组合字符串,多个记录之间通过换行符链接。
    第一个字段需要根据NeedPrep和Prepared来返回结果:
        if not NeedPrep then
            return ChrW(&H2605) '实心五角星
        else
            if Prepared then
                return ChrW(&H2022) '实心圆点
            else
                return ChrW(&H25EF) '空心圆
            end if
        end if
    Pattern为数字加"|"的字符串,用以定义各字段的长度。如果数字部分小于对应字段长度,则直接返回字段,如果数字部分大于字段长度,则将字段值通过空格补齐。
    举例: Pattern为5|8|0,表示这个pattern中包含3个字段,对应了要求中的NeedPrep/Prepared, Name 和Description。

### 新增辅助函数与过程

#### 判断单元格是否有名称
```vba
Public Function HasCellName(cell As Range) As Boolean
    On Error Resume Next
    Dim n As String
    n = cell.Name.NameLocal
    HasCellName = (Err.Number = 0)
    On Error GoTo 0
End Function
```

#### 批量清理数据块的主过程
```vba
Public Sub PrepDataBlocksBetweenNames(ByVal StartName As String, Optional HasHeadLine As Boolean = False)
    Dim startCell As Range
    ' 1. 定位起始单元格
    On Error Resume Next
    Set startCell = shGeneral.Range(StartName)
    On Error GoTo 0
    If startCell Is Nothing Then
        MsgBox "Named range '" & StartName & "' not found.", vbExclamation, "Error"
        Exit Sub
    End If

    Dim startRow As Long
    startRow = startCell.Row
    Dim col As Long
    col = startCell.Column

    ' 2. 向下查找下一个有名称的单元格，最多查找100行
    Dim endRow As Long
    Dim i As Long
    endRow = startRow + 100
    For i = startRow + 1 To startRow + 100
        If i > shGeneral.Rows.Count Then Exit For
        If HasCellName(shGeneral.Cells(i, col)) Then
            If shGeneral.Cells(i, col).Name.NameLocal <> shGeneral.Cells(i, col).Address(False, False, xlA1, True) Then
                endRow = i
                Exit For
            End If
        End If
    Next i
    If endRow > shGeneral.Rows.Count Then endRow = shGeneral.Rows.Count

    ' 3. 获取区域（开始行+1 到 结束行-1）
    Dim regionStart As Long, regionEnd As Long
    regionStart = startRow + 1
    regionEnd = endRow - 1
    If regionStart > regionEnd Then Exit Sub ' 区域无效

    ' 只遍历已用区域的列
    Dim usedColCount As Long
    usedColCount = shGeneral.UsedRange.Columns.Count
    Dim r As Long, c As Long
    For r = regionStart To regionEnd
        For c = 1 To usedColCount
            Dim cell As Range
            Set cell = shGeneral.Cells(r, c)
            If HasCellName(cell) Then
                If cell.Name.NameLocal <> cell.Address(False, False, xlA1, True) Then
                    Call PrepDataBlockByRange(cell.Name.NameLocal)
                End If
            End If
        Next c
    Next r
End Sub
```

### 技术细节
1. 主要内容都在本项目的.\CharacterManagementTool.xlsb文件中维护,包括表格,数据,代码.本文档以后都称这个文件为"Excel文件".
2. 在VBA窗口中将一些页面的CodeName改名,方便在VBA中直接作为WorkSheet对象调用. 具体列表参考后续的[[# Excel文件页面列表]]
3. 在Excel中的shGeneral为交互主界面,上面定义一系列的名称Names用以传递当前Character的数值






### Excel文件页面列表
本段落为Excel文件中所有使用到的特定页面,每个页面显示CodeName - Name, 后续再加上细节描述.

1. shTableSchema - Table Schema
    该页面包含所有需要用到的表格的结构. 因为Cursor无法直接读取Excel文件内容,此处内容不再展开.具体表格设计在其他部分介绍.
2. shGeneral - General
    该页面为主页面,需要进行一定的美化.用户可以在这个页面选择新增/修改/删除角色.并且在新增/修改角色的时候对角色各项内容进行修改.


### 表格细节
本段落为Excel中包含的表格以及其字段明细
1. CharacterMasterSchema

用于存储角色主要信息

|字段|类型|名字|打印输出|备注|
|---|---|---|---|---|
|CharacterID|Int|角色ID|N|用于在表格间关联,界面上不显示|
|CharacterType|String|角色类型|N|Player, NPC|
|CharacterStatus|String|角色状态|N|Alive,Dead|
|Player|String|玩家名|Y||
|Character|String|角色名|Y||
|Background|String|背景|Y||
|Class|String|职业|Y||
|ClassLv|Int|职业等级|Y||
|Race|String|种族|Y||
|Alignment|String|阵营|Y||
|Exp|long|经验值|N||
|Strength|Int|力量|Y||
|StrengthAdd|Int|力量加值|Y||
|Dexterity|Int|敏捷|Y||
|DexterityAdd|Int|敏捷加值|Y||
|Constitution|Int|体质|Y||
|ConstitutionAdd|Int|体质加值|Y||
|Intelligence|Int|智力|Y||
|IntelligenceAdd|Int|智力加值|Y||
|Wisdom|Int|感知|Y||
|WisdomAdd|Int|感知加值|Y||
|Charisma|Int|魅力|Y||
|CharismaAdd|Int|魅力加值|Y||
|ArmorClass|Int|护甲等级|Y||
|Initiative|Int|先攻|Y||
|Speed|Int|速度|Y||
|Inspiration|Int|激励|N||
|ProficiencyBonus|Int|熟练加值|Y||
|SavingThrowStr|Int|力量豁免|Y||
|SavingThrowDex|Int|敏捷豁免|Y||
|SavingThrowCon|Int|体质豁免|Y||
|SavingThrowInt|Int|智力豁免|Y||
|SavingThrowWis|Int|感知豁免|Y||
|SavingThrowCha|Int|魅力豁免|Y||
|SavingThrowStrP|Bool|力量豁免熟练|Y|是否熟练项Proficiency|
|SavingThrowDexP|Bool|敏捷豁免熟练|Y|是否熟练项Proficiency|
|SavingThrowConP|Bool|体质豁免熟练|Y|是否熟练项Proficiency|
|SavingThrowIntP|Bool|智力豁免熟练|Y|是否熟练项Proficiency|
|SavingThrowWisP|Bool|感知豁免熟练|Y|是否熟练项Proficiency|
|SavingThrowChaP|Bool|魅力豁免熟练|Y|是否熟练项Proficiency|
|SkillAcrobatics|Int|技能体操(敏捷)|Y||
|SkillAnimalHandling|Int|技能驯养动物(感知)|Y||
|SkillArcana|Int|技能奥秘知识(智力)|Y||
|SkillAthletics|Int|技能运动(力量)|Y||
|SkillDeception|Int|技能欺瞒(魅力)|Y||
|SkillHistory|Int|技能历史知识(智力)|Y||
|SkillInsight|Int|技能洞察(感知)|Y||
|SkillIntimidation|Int|技能威吓(魅力)|Y||
|SkillInvestigation|Int|技能调查(智力)|Y||
|SkillMedicine|Int|技能医药(感知)|Y||
|SkillNature|Int|技能自然知识(智力)|Y||
|SkillPerception|Int|技能察觉(感知)|Y||
|SkillPerformance|Int|技能表演(魅力)|Y||
|SkillPersuasion|Int|技能说服(魅力)|Y||
|SkillReligion|Int|技能宗教知识(智力)|Y||
|SkillSleightOfHand|Int|技能手上功夫(敏捷)|Y||
|SkillStealth|Int|技能隐匿(敏捷)|Y||
|SkillSurvival|Int|技能生存(感知)|Y||
|SkillAcrobaticsP|Bool|技能熟练体操(敏捷)|Y|是否熟练项Proficiency|
|SkillAnimalHandlingP|Bool|技能熟练驯养动物(感知)|Y|是否熟练项Proficiency|
|SkillArcanaP|Bool|技能熟练奥秘知识(智力)|Y|是否熟练项Proficiency|
|SkillAthleticsP|Bool|技能熟练运动(力量)|Y|是否熟练项Proficiency|
|SkillDeceptionP|Bool|技能熟练欺瞒(魅力)|Y|是否熟练项Proficiency|
|SkillHistoryP|Bool|技能熟练历史知识(智力)|Y|是否熟练项Proficiency|
|SkillInsightP|Bool|技能熟练洞察(感知)|Y|是否熟练项Proficiency|
|SkillIntimidationP|Bool|技能熟练威吓(魅力)|Y|是否熟练项Proficiency|
|SkillInvestigationP|Bool|技能熟练调查(智力)|Y|是否熟练项Proficiency|
|SkillMedicineP|Bool|技能熟练医药(感知)|Y|是否熟练项Proficiency|
|SkillNatureP|Bool|技能熟练自然知识(智力)|Y|是否熟练项Proficiency|
|SkillPerceptionP|Bool|技能熟练察觉(感知)|Y|是否熟练项Proficiency|
|SkillPerformanceP|Bool|技能熟练表演(魅力)|Y|是否熟练项Proficiency|
|SkillPersuasionP|Bool|技能熟练说服(魅力)|Y|是否熟练项Proficiency|
|SkillReligionP|Bool|技能熟练宗教知识(智力)|Y|是否熟练项Proficiency|
|SkillSleightOfHandP|Bool|技能熟练手上功夫(敏捷)|Y|是否熟练项Proficiency|
|SkillStealthP|Bool|技能熟练隐匿(敏捷)|Y|是否熟练项Proficiency|
|SkillSurvivalP|Bool|技能熟练生存(感知)|Y|是否熟练项Proficiency|
|PassiveWisdom|Int|被动感知(察觉)|Y||
|MaxHP|Int|最大生命值|Y||
|CurHP|Int|当前生命值|Y||
|TmpHP|Int|临时生命值|Y||
|HD|Int|生命骰|Y|显示为D8,D10之类|
|MaxHD|Int|生命骰数|Y|结合上面显示为MaxHD&D&HD, 举例4D8|
|MoneyCP|Int|铜币|Y||
|MoneySP|Int|银币|Y||
|MoneyEP|Int|金银币|Y||
|MoneyGP|Int|金币|Y||
|MoneyPP|Int|铂金币|Y||
|Age|Int|年龄|Y||
|Height|String|身高|Y||
|Weight|String|体重|Y||
|Eyes|String|瞳色|Y||
|Skin|String|肤色|Y||
|Hair|String|发色|Y||
|SpellCastingClass|String|施法职业|Y||
|SpellCastingAbility|String|施法关键属性|Y||
|SpellSaveDC|Int|法术豁免DC|Y||
|SpellAttackBouns|Int|法术攻击加值|Y||


2. CharacterMemoSchema

用于存储角色其他备注信息

|字段|类型|名字|打印输出|备注|
|---|---|---|---|---|
|CharacterID|Int|角色ID|N|用于在表格间关联,界面上不显示|
|MemoType|String|备注类型|N|指向名称MemoTypeList|
|Contents|String|备注内容|Y||

注意:
MemoType的值以及对应的含义:
PersonalityTraits       个人特点
Ideals                  理想
Bonds                   羁绊
Flaws                   缺点
OtherProficiencies      其他熟练项
Languages               语言
Features                特性
Traits                  专长
CharacterAppearance     角色形象
Allies                  同盟
Organizations           组织
AdditionalFeatures      额外特性
AdditionalTraits        额外专长
CharacterBackstory      背景故事
Treasure                财物

3. CharacterAttackSpellSchema

用于存储角色武器与法术列表

|字段|类型|名字|打印输出|备注|
|---|---|---|---|---|
|CharacterID|Int|角色ID|N|用于在表格间关联,界面上不显示|
|Type|String|物品类型|N|指向名称ItemTypeList|
|Name|String|名称|Y||
|AtkBonus|Int|攻击加值|Y||
|Damage_Type|String|伤害与类型|Y||
|SpellMemo|String|法术备注|Y||
|Attuned|Bool|是否同调|Y|如果武器已同调,则在武器名前打印*号|
|Equiped|Bool|是否装备|Y|如果武器已装备则打印,否则在装备中列举|

4. CharacterEquipmentSchema

用于存储角色物品列表

|字段|类型|名字|打印输出|备注|
|---|---|---|---|---|
|CharacterID|Int|角色ID|N|用于在表格间关联,界面上不显示|
|Type|String|物品类型|N|指向名称ItemTypeList|
|Name|String|名称|||
|Quantity|String|数量|||
|Attuned|Bool|是否同调|Y|如果装备已同调,则在装备名前打印*号|
|Equiped|Bool|是否装备|Y|如果已装备则优先打印,否则在下方打印|

5. CharacterSpellSchema

用于存放角色法表

|字段|类型|名字|打印输出|备注|
|---|---|---|---|---|
|CharacterID|Int|角色ID|N|用于在表格间关联,界面上不显示|
|SpellLevel|Int|法术等级|N|法术等级,从0到9, 0为戏法|
|Name|String|法术名称|Y||
|Description|String|法术描述|Y||
|Prepared|Bool|是否准备|Y||

6. CharacterSpellSlotSchema

用于存放角色法术位列表

|字段|类型|名字|打印输出|备注|
|---|---|---|---|---|
|CharacterID|Int|角色ID|N|用于在表格间关联,界面上不显示|
|SpellLevel|Int|法术等级|N|法术等级,从0到9, 0为戏法|
|SlotsTotal|Int|法术位|Y||
|SlotsExpended|Int|已使用法术位|Y||

## Code Implementation

### CharacterToUI Procedure
```vba
Option Explicit

' Constants for property types
Private Const PROP_TYPE_COLLECTION As String = "Collection"

Public Sub CharacterToUI(ByVal CharacterID As Long)
    ' Check if character exists in dictionary
    If Not Characters.Exists(CharacterID) Then
        MsgBox "Character ID " & CharacterID & " not found in dictionary", vbExclamation, "Error"
        Exit Sub
    End If
    
    ' Get character from dictionary
    Dim Character As CharacterMaster
    Set Character = Characters(CharacterID)
    
    ' Get all properties from CharacterMaster class
    Dim Prop As Variant
    Dim PropValue As Variant
    Dim RangeExists As Boolean
    Dim Response As VbMsgBoxResult
    
    ' Get properties from CharacterMasterSchema table
    Dim Properties As Variant
    Properties = GetPropertiesFromSchema("CharacterMasterSchema")
    
    For Each Prop In Properties
        ' Skip Collection type properties
        PropValue = CallByName(Character, Prop, VbGet)
        If TypeName(PropValue) = PROP_TYPE_COLLECTION Then
            ' Skip this property as it's a collection
            ' Collections will be handled in future updates
        Else
            ' Check if named range exists in shGeneral
            On Error Resume Next
            RangeExists = Not shGeneral.Range(Prop) Is Nothing
            On Error GoTo 0
            
            If Not RangeExists Then
                ' Ask user if should continue
                Response = MsgBox("Member " & Prop & " not found in UI. Continue writing other members?", _
                                vbQuestion + vbYesNo, "Missing Member")
                
                If Response = vbNo Then
                    Exit Sub
                End If
            Else
                ' Write property value to UI
                shGeneral.Range(Prop) = PropValue
            End If
        End If
    Next Prop
End Sub

' Helper function to get properties from schema table
Private Function GetPropertiesFromSchema(ByVal SchemaName As String) As Variant
    Dim Properties As Collection
    Set Properties = New Collection
    
    ' Get the schema table
    Dim SchemaTable As ListObject
    Set SchemaTable = shTableSchema.ListObjects(SchemaName)
    
    ' Get the field names from the schema
    Dim DataRange As Range
    Set DataRange = SchemaTable.ListColumns("字段").DataBodyRange
    
    Dim Cell As Range
    For Each Cell In DataRange
        If Not IsEmpty(Cell) Then
            Properties.Add Cell.Value
        End If
    Next Cell
    
    ' Convert collection to array
    Dim Result() As Variant
    ReDim Result(1 To Properties.Count)
    Dim i As Long
    For i = 1 To Properties.Count
        Result(i) = Properties(i)
    Next i
    
    GetPropertiesFromSchema = Result
End Function
```

This implementation:
1. Uses the schema table to dynamically get properties
2. Maintains all the required functionality
3. Is more maintainable as it automatically adapts to schema changes
4. Follows VBA coding standards with proper variable declarations

Key improvements:
1. Removed hardcoded property list
2. Added GetPropertiesFromSchema function to read properties from schema table
3. Added constant for property type checking
4. More maintainable and flexible design

Benefits of this approach:
1. Automatically adapts to schema changes
2. No need to manually update property lists
3. Reduces chance of errors
4. Single source of truth (schema table)
5. Easier to maintain

### PrepareCharacter Procedure
```vba
' Helper function to get maximum ID from dictionary
Private Function GetMaxCharacterId() As Long
    Dim MaxId As Long
    MaxId = 0
    
    Dim Key As Variant
    For Each Key In Characters.Keys
        If CLng(Key) > MaxId Then
            MaxId = CLng(Key)
        End If
    Next Key
    
    GetMaxCharacterId = MaxId
End Function

Public Sub PrepareCharacter(ByVal CharacterId As Variant)
    ' Initialize objects and read data
    Initialize
    ReadCharacters
    
    ' Check if CharacterId is numeric
    If IsNumeric(CharacterId) Then
        ' Convert to Long for dictionary lookup
        Dim CharacterIdLong As Long
        CharacterIdLong = CLng(CharacterId)
        
        ' Check if character exists in dictionary
        If Characters.Exists(CharacterIdLong) Then
            ' Character exists, update UI
            CharacterToUI CharacterIdLong
        Else
            ' Character does not exist
            MsgBox "Character ID " & CharacterId & " does not exist", vbExclamation, "Error"
        End If
    Else
        ' New character - get max ID and add 1
        Dim NewCharacterId As Long
        NewCharacterId = GetMaxCharacterId() + 1
        
        ' Create new character and add to dictionary
        Dim NewCharacter As CharacterMaster
        Set NewCharacter = New CharacterMaster
        NewCharacter.CharacterId = NewCharacterId
        
        ' Add to dictionary
        Characters.Add NewCharacterId, NewCharacter
        
        ' Update UI with new character
        CharacterToUI NewCharacterId
    End If
    
    ' Terminate to clean up objects
    Terminate
End Sub
```

This implementation:
1. Adds a helper function GetMaxCharacterId to encapsulate the max ID logic
2. Makes the code more modular and easier to maintain
3. Improves code readability
4. Makes the max ID logic reusable if needed elsewhere
5. Maintains all existing functionality

Key features:
1. Separates concerns by moving max ID logic to a dedicated function
2. Makes the main procedure cleaner and more focused
3. Follows the DRY (Don't Repeat Yourself) principle
4. Maintains all existing functionality
5. Improves code organization

### 事件处理
1. Worksheet_Change 事件
    用于监听工作表单元格值的变化
    - 在需要监听的工作表模块中实现
    - 语法如下:
    ```vba
    Private Sub Worksheet_Change(ByVal Target As Range)
        ' 检查是否是单个单元格发生变化
        If Target.Cells.Count > 1 Then
            Exit Sub
        End If
        
        ' 检查是否是目标单元格发生变化
        If Not Intersect(Target, Range("CharacterIDName")) Is Nothing Then
            ' 暂时禁用事件以防止循环触发
            Application.EnableEvents = False
            
            ' 执行你的代码
            ' 例如: PrepareCharacter Target.Value
            
            ' 恢复事件
            Application.EnableEvents = True
        End If
    End Sub
    ```
    - 注意事项:
        1. 必须放在工作表模块中(例如 shGeneral 的代码模块)
        2. 使用名称引用而不是硬编码的单元格地址,提高代码的可维护性
        3. 使用 Target.Cells.Count 检查是否只有单个单元格发生变化
        4. 使用 Application.EnableEvents 控制事件触发,防止循环调用
        5. 如果监听多个命名单元格,可以使用 Union 函数组合多个 Range
        6. 建议在事件处理代码中添加错误处理,确保 Application.EnableEvents 被正确恢复
```

## 新增过程

### Public Sub SaveCurrentCharacter

该过程用于将当前 shGeneral 页面上的角色信息保存到数据表：

```vba
Public Sub SaveCurrentCharacter()
    ' Initialize objects and read data
    Initialize
    ReadCharacters
    
    ' Get CharacterID from shGeneral
    Dim CharacterID As Variant
    CharacterID = shGeneral.Range("CharacterID").Value
    
    ' Ensure CharacterID is numeric and not empty
    If IsNumeric(CharacterID) And Not IsEmpty(CharacterID) Then
        Dim CharacterIDLong As Long
        CharacterIDLong = CLng(CharacterID)
        
        ' Read character from UI
        Dim newCharacter As CharacterMaster
        Set newCharacter = UIToCharacter(CharacterIDLong)
        
        ' Update or add to Characters dictionary
        If Characters.Exists(CharacterIDLong) Then
            Set Characters(CharacterIDLong) = newCharacter
        Else
            Characters.Add CharacterIDLong, newCharacter
        End If
    Else
        MsgBox "Invalid CharacterID in shGeneral!", vbExclamation, "Error"
        Terminate
        Exit Sub
    End If
    
    ' Write all characters to sheets
    WriteCharacters
    
    ' Terminate to clean up objects
    Terminate
End Sub
```

**说明：**
- 先初始化并读取所有数据。
- 读取 shGeneral 页的 CharacterID。
- 调用 UIToCharacter 读取当前 UI 上的角色信息。
- 如果字典中已存在该角色，则替换，否则新增。
- 最后写回所有角色并清理对象。
```

### Public Function CreateWordFromTemplate(ByVal TemplateName As String) As Object

该函数用于根据指定模板名称，从Note页命名单元格获取Word模板文件名，并在\Documents\目录下以该模板新建Word文档，返回新建文档对象。

```vba
Public Function CreateWordFromTemplate(ByVal TemplateName As String) As Object
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim templatePath As String, fileName As String
    Dim docFolder As String
    
    ' 1. 读取Word文件名（从shNote或Worksheets("Note")）
    On Error Resume Next
    fileName = shNote.Range(TemplateName).Value
    If fileName = "" Then
        fileName = Worksheets("Note").Range(TemplateName).Value
    End If
    On Error GoTo 0
    If fileName = "" Then
        MsgBox "未指定Word文件名！", vbExclamation
        Set CreateWordFromTemplate = Nothing
        Exit Function
    End If
    
    ' 2. 构造完整路径
    docFolder = ThisWorkbook.Path & "\Documents\"
    If Right(docFolder, 1) <> "\" Then docFolder = docFolder & "\"
    templatePath = docFolder & fileName
    
    If Dir(templatePath) = "" Then
        MsgBox "找不到模板文件：" & templatePath, vbExclamation
        Set CreateWordFromTemplate = Nothing
        Exit Function
    End If
    
    ' 3. 用模板新建文档
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = True
    Set wordDoc = wordApp.Documents.Add(Template:=templatePath, NewTemplate:=False)
    
    ' 4. 返回文档对象
    Set CreateWordFromTemplate = wordDoc
End Function
```

**功能说明：**
- 通过TemplateName参数，查找Note页对应命名单元格，获取Word模板文件名。
- 在\Documents\目录下查找该模板文件。
- 用该模板新建Word文档，返回文档对象（wordDoc）。
- 若找不到文件或名称，弹出消息框并返回Nothing。

**注意事项：**
- 返回的wordDoc对象由调用者负责后续操作和释放（如填充内容、保存、关闭、Set为Nothing等）。
- 不要在函数内部将wordDoc设为Nothing，否则外部无法继续操作该文档。

### Public Sub ExportCharacter()

该过程用于将当前shGeneral页面上的角色信息导出到Word文档。

```vba
Public Sub ExportCharacter()
    Dim CharacterID As Variant
    Dim Character As CharacterMaster
    Dim wordDoc As Object
    
    ' 1. 获取当前角色ID
    CharacterID = shGeneral.Range("CharacterID").Value
    If Not IsNumeric(CharacterID) Or IsEmpty(CharacterID) Then
        MsgBox "无效的角色ID！", vbExclamation
        Exit Sub
    End If
    
    ' 2. 获取当前角色对象
    Set Character = UIToCharacter(CLng(CharacterID))
    If Character Is Nothing Then
        MsgBox "未能获取角色对象！", vbExclamation
        Exit Sub
    End If
    
    ' 3. 创建Word文档（模板名可根据实际情况调整）
    Set wordDoc = CreateWordFromTemplate("CharacterSheet")
    If wordDoc Is Nothing Then
        MsgBox "Word文档创建失败！", vbExclamation
        Exit Sub
    End If
    
    ' 4. 写入角色名到Word文档Tag为"Character"的ContentControl
    Call PrintToWord(wordDoc, "Character", Character.Character)
    
    ' TODO: 后续可补充写入更多内容
    MsgBox "角色导出完成，后续内容请补充PrintToWord调用。", vbInformation
End Sub
```

**功能说明：**
- 获取当前shGeneral页面的CharacterID。
- 调用UIToCharacter获取当前角色对象。
- 调用CreateWordFromTemplate生成Word文档（模板名为"CharacterSheet"，可根据实际情况调整）。
- 调用PrintToWord，将角色名写入Word文档中Tag为"Character"的ContentControl。
- 预留后续可扩展写入更多内容。

**注意事项：**
- 需保证Word模板中存在Tag为"Character"的ContentControl。
- 后续如需批量写入更多内容，只需补充多次PrintToWord调用。

## Word自动化开发补充说明

在本项目中，针对Excel与Word自动化联动，实际开发和测试过程中遇到如下关键问题点和注意事项，建议后续开发时务必参考：

1. **通过模板方式新建Word文档**
   - 推荐用 `Documents.Add(Template:=...)` 以模板（可为 `.docx` 文件）新建文档，而不是复制文件或直接打开。
   - 优点：不会影响原模板文件，生成的新文档初始为"未保存"状态，便于后续保存为任意文件名。

2. **内容控件（ContentControl）的遍历**
   - 内容控件可能存在于多个区域：主文档区（正文）、文本框（Shape/TextBox）、页眉、页脚、表格等。
   - 遍历方式：
     - 主文档区：`For Each cc In wordDoc.ContentControls`
     - 文本框等Shape中：`For Each shp In wordDoc.Shapes`，再 `For Each cc In shp.TextFrame.TextRange.ContentControls`
   - 只遍历 `wordDoc.ContentControls` 不能覆盖所有情况，必须同时遍历 `Shapes`。

3. **For Each 中变量类型声明**
   - 必须用 `Dim cc As Object` 或 `Dim cc As Word.ContentControl`，不能用 `Dim cc As Shape` 或其他类型，否则会"类型不匹配"。
   - 同理，`shp` 也建议用 `Dim shp As Object`，兼容性更好。

4. **内容控件的Tag用法**
   - 推荐用Tag作为唯一标识，便于VBA自动化查找和赋值。
   - 设计模板时，所有需要自动化填充的内容控件都应设置唯一Tag。

5. **对象释放与资源管理**
   - 即使不关闭Word窗口，也应 `Set wordDoc = Nothing`，`Set wordApp = Nothing`，防止内存泄漏。

6. **模板文件格式**
   - `.docx` 也可作为模板，无需专门保存为 `.dotx`。
   - 用 `Documents.Add(Template:=...)` 创建的新文档不会影响原文件，也不会自动保存。

7. **内容控件在文本框/Shape中的访问**
   - VBA只能通过 `Shape.TextFrame.TextRange.ContentControls` 访问文本框中的内容控件，不能通过主文档集合访问。

8. **其他建议**
   - 若有文本框、页眉页脚等特殊区域，务必在VBA中遍历所有相关集合。
   - 测试时可先用 `.Visible = True`，正式批量生成时可设为 `False`。

如需将本补充说明扩展为开发文档、代码注释模板或进一步细化，请随时补充！

### Public Function PrintSignedNumber(ByVal num As Integer) As String

该函数用于将整数以带符号的字符串形式输出，常用于属性加值等场景。

```vba
Public Function PrintSignedNumber(ByVal num As Integer) As String
    If num > 0 Then
        PrintSignedNumber = "+" & CStr(num)
    ElseIf num < 0 Then
        PrintSignedNumber = CStr(num)
    Else
        PrintSignedNumber = ""
    End If
End Function
```

**功能说明：**
- num > 0 时返回 "+5" 形式
- num < 0 时返回 "-4" 形式
- num = 0 时返回空字符串

**示例：**
- PrintSignedNumber(5)   → "+5"
- PrintSignedNumber(-4)  → "-4"
- PrintSignedNumber(0)   → ""

---

### Public Function PrintBoolean(ByVal val As Boolean) As String

该函数用于将布尔值以符号形式输出，常用于Word/Excel等文本输出格式化。

```vba
Public Function PrintBoolean(ByVal val As Boolean) As String
    If val = True Then
        PrintBoolean = ChrW(&H2022)
    Else
        PrintBoolean = ""
    End If
End Function
```

**功能说明：**
- val = True 时返回黑点符号（ChrW(&H2022)）
- val = False 时返回空字符串

**示例：**
- PrintBoolean(True)  → "•"
- PrintBoolean(False) → ""
```

### Public Function ExportValues(ByVal SplitString As String, ParamArray PropNames() As Variant) As String

该方法用于按顺序拼接指定属性值，中间用SplitString分隔。

```vba
ExportValues(" Lv ", "Class", "ClassLv") ' 结果如："战士 Lv 5"
```

- 依次读取参数列表中的属性名，取当前对象的属性值，按SplitString连接。

---

### Public Function ExportMemoLists(ParamArray MemoTypes() As Variant) As String

该方法用于导出CharacterMemoList中指定类型的内容，内容间用换行符分隔。

```vba
ExportMemoLists("Features", "Traits")
' 结果为所有MemoType为Features或Traits的Contents内容，按行拼接
```

- 遍历CharacterMemoList，匹配MemoType，拼接Contents。

---

### Public Function ExportAttackLists(ByVal Pattern As String, ByVal Equiped As Boolean) As String

该方法用于导出装备状态为Equiped的攻击列表，按Pattern格式化。

```vba
ExportAttackLists("0|12|5|8|0", True)
' 结果如：
' •阳炎剑         (*)  +8      1d8+2,光耀
' •吸失盾                      AC+2
```

- Pattern为数字加|的字符串，定义各字段宽度。
- 字段顺序：Equiped, Name, Attuned, AtkBonus, Damage_Type。
- Equiped为True输出黑点，Attuned为True输出(*)，AtkBonus用PrintSignedNumber格式化。
- 多条记录用换行符分隔。

---

### Public Function ExportAtkSpellLists(ByVal Pattern As String, ByVal Equiped As Boolean) As String

该方法用于导出装备状态为Equiped的法术列表，按Pattern格式化。

```vba
ExportAtkSpellLists("0|12|16|0", True)
' 结果如：
' •龙息          2d6,寒冷          种族:向前喷吐一道15英尺锥形的吐息.(短/长休一次)DC12+体质
```

- Pattern为数字加|的字符串，定义各字段宽度。
- 字段顺序：Equiped, Name, Damage_Type, SpellMemo。
- Equiped为True输出黑点。
- 多条记录用换行符分隔。

---

### Public Function ExportEquipmntList(ByVal Pattern As String) As String

该方法用于导出装备列表，按Pattern格式化。

```vba
ExportEquipmntList("8|0|0")
' 结果如：
' 龙鳞甲     (*)1
' 长剑      1
```

- Pattern为数字加|的字符串，定义各字段宽度。
- 字段顺序：Name, Attuned, Quantity。
- Attuned为True输出(*)。
- 多条记录用换行符分隔。
```

### Public Sub FastFillWordContentControls(wordDoc As Object, dict As Object)

该过程用于将字典中的内容批量写入Word文档的内容控件中。

```vba
Sub FastFillWordContentControls(wordDoc As Object, dict As Object)
    Dim wordApp As Object
    Set wordApp = wordDoc.Application
    wordApp.ScreenUpdating = False
    wordApp.EnableEvents = False

    Dim cc As Object, shp As Object
    ' 正文
    For Each cc In wordDoc.ContentControls
        If dict.Exists(cc.Tag) Then
            cc.Range.Text = dict(cc.Tag)
        End If
    Next
    ' Shapes
    For Each shp In wordDoc.Shapes
        If shp.Type = 17 Then
            For Each cc In shp.TextFrame.TextRange.ContentControls
                If dict.Exists(cc.Tag) Then
                    cc.Range.Text = dict(cc.Tag)
                End If
            Next
        End If
    Next

    wordApp.ScreenUpdating = True
    wordApp.EnableEvents = True
End Sub
```

**功能说明：**
- 该过程用于将字典中的内容批量写入Word文档的内容控件中。
- 通过遍历Word文档中的所有内容控件，将字典中的内容写入相应的控件中。
- 该过程可以提高写入效率，避免多次遍历控件。

**注意事项：**
- 该过程需要传入一个Word文档对象和一个字典对象。
- 字典中的键值对需要与Word文档中的ContentControl的Tag相对应。
- 该过程会关闭屏幕刷新和事件，以提高写入效率。

### 文本框内容自适应字号（防止内容溢出）

**需求描述：**
- 在Word模板中，若文本框（Shape）内的富文本ContentControl写入内容过长，可能会被截断。
- 希望自动检测内容是否溢出，并在溢出时递减字号，直至内容完整显示或达到最小字号。
- 不希望在代码中指定初始字号，而是以模板中设置的字号为起点。

**推荐实现思路：**
- Word VBA 没有直接的 BoundWidth/BoundHeight 属性用于检测文本溢出。
- 可采用"递减字号，直到文本框内容和原内容一致（未被截断）"的变通法。
- 适用于普通文本框+富文本内容控件的场景。

**示例代码：**

```vba
Sub FitContentControlFontToTextbox(cc As Object, shp As Object, Optional minFontSize As Integer = 8)
    Dim curFontSize As Integer
    Dim originalText As String
    curFontSize = cc.Range.Font.Size
    If curFontSize = 0 Then curFontSize = 12 ' 兜底
    originalText = cc.Range.Text

    ' 递减字号直到文本框内容和原内容一致（未被截断），或到达最小字号
    Do While (shp.TextFrame.TextRange.Text <> originalText) And curFontSize > minFontSize
        curFontSize = curFontSize - 1
        cc.Range.Font.Size = curFontSize
        DoEvents
    Loop
End Sub
```

**用法说明：**
- 先写入内容，再调用该过程。
- 只在内容溢出时递减字号，起点为模板中ContentControl的字号。
- 可批量处理所有文本框中的ContentControl。

**注意事项：**
- 该方法为变通法，适用于大多数普通文本框场景。
- 若内容控件内有复杂格式或特殊换行，需进一步定制。
- Word VBA 不支持 BoundWidth/BoundHeight，不能像PowerPoint那样直接检测溢出。