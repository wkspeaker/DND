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
    在Terminate过程之前执行,或者在专门的Save命令中调用.用于将当前的Characters对象保存到相应表格中
    - 首先将关联到的数据表的记录清空,比如shCharacterMaster.ListObject(1).DataBodyRange清空
    - 针对数据字典Characters中的每一个项目,读取相应的CharacterMaster对象,将其中的各个Collection中的记录同样遍历,写入各个数据表中.
        - Collection CharacterMemoList 写入 shCharacterMemo.ListObject(1)
        - Collection CharacterAttackSpellList 写入shCharacterAttackSpell.ListObject(1)
        - Collection CharacterEquipmentList 写入shCharacterEquipment.ListObject(1)
        - CharacterMaster对象本身的其他属性写入shCharacterMaster.ListObject(1)
7. Public Sub CharacterToUI(ByVal CharacterID)
    在ReadCharacters过程执行以后运行该过程,用于将数据字典中的Character的各项数值写到shGeneral页面中的相应字段
    因此该过程运行的前提是已经将各个数据表的内容读取完毕,并且生成了各个Character的数据字典
    传入参数为CharacterID,从数据字典Characters中定位到对应的角色,然后写入shGeneral页面相应的位置:
    - 目前该功能未完善,先处理类CharacterMaster中的各个成员,对于包含从表信息的各个Collection暂时不考虑,以后再追加
    - 对于类中除了Collection的各个成员,已经在shGeneral页面中定义了相应的名称, 将各个值写入就可以
        举例: Character.Strength = 6, 则通过shGeneral.Range("Strength") = Character.Strength来完成写入
        假设条件是类中的属性成员名字一定和shGeneral中的名称定义吻合
        需要考虑到两者有差异的情况. 如果发现有类中成员没有在shGeneral中找到对应名称的情况,弹出对话框显示"成员Strength没有找到对应名称,是否继续写入?", 是,否
        如果用户选择是,则忽略当前成员继续完成剩余部分的写入,如果选择否,则终止当前过程.
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