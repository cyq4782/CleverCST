' GNU GENERAL PUBLIC LICENSE
'                      Version 3, 29 June 2007
' Copyright (C) 2007 Free Software Foundation, Inc. <https://fsf.org/>
' Everyone is permitted to copy and distribute verbatim copies
' of this license document, but changing it is not allowed.

Option Explicit
' 常量定义
Public Const ConfigParNumber = 6 ' 需要配置的遗传算法参数总数
Public Const ParNumberMax = 100 ' 最多允许的待优化参数数目
Public Const TurnNumberMax = 1000 ' 遗传代数的最大值
Public Const TempDir = ""  ' 临时文件存放目录，形如：E:\qtemp\
Public Const Mon1D = "1D Results\S-Parameters\S1,1[1.0,0.0]+2[1.0,90],[10]" ' 1D监视器名称
Public Const Mon1Dp = "1D Results\S-Parameters\S2,1[1.0,0.0]+2[1.0,90],[10]" ' 另一个1D监视器名称
Public Const Mon2D = "2D/3D Results\E-Field\e-field (f=f) [1[1.0,0.0]+2[1.0,90],[10]]" ' 2D监视器名称
' 完成常量定义

' 定义ConfigType类型对配置进行存储
Type ConfigType
    par_num As Long ' 待优化参数个数
    n_live As Long ' 每一代存活的个体数量
    n_all As Long ' 每一代的个体总数
    turns As Long ' 求解代数
    P_BY As Double ' 发生变异的概率
    P_AV As Double ' 发生平均化遗传的概率
End Type
' 完成ConfigType类型的定义

' 定义ParType类型对待优化参数进行存储
Type ParType
    NameStr As String ' 参数名
    MinNum As Double ' 参数最小值
    MaxNum As Double ' 参数最大值
    StepNum As Double ' 参数步长
End Type
' 完成ParType类型的定义

' 定义BodyType类型对个体进行存储
Type BodyType
    ParValue(1 To ParNumberMax) As Double ' 按照ParList顺序存储的个体参数值
    GoodValue As Double ' 计算得到的适应度
End Type
' 完成BodyType类型的定义



' 定义公共变量
Public MyConfig As ConfigType ' MyConfig存储关于遗传算法的配置信息
Public ParList(1 To ParNumberMax) As ParType ' ParList存储待优化的参数列表
Public BodyList(1 To 1000) As BodyType ' BodyList存储种群个体列表
Public BodyLiveList(1 To 1000) As BodyType ' BodyLiveList存储存活个体列表
Public TurnAlready As Long ' TurnAlready存储遗传算法的轮数
Public BestBodyList(1 To TurnNumberMax) As BodyType ' BestBodyList存储每一代最优解个体列表
'--------------------------------------------
Sub Main()
Dim a As Double
Dim body As BodyType


ReadConfig "E:\CYQ\遗传算法测试\config.txt"
ReadParConfig "E:\CYQ\遗传算法测试\par_config.txt"
DoGA

End Sub
'--------------------------------------------

Public Sub ReadConfig(filename As String)
' 此过程用于读取配置文件，唯一参数FileName指示配置文件路径，配置文件应具有以下格式：
'------------------------------------
' par_num=3
' n_live=10
' n_all=100
' turns=50
' P_BY=10
' P_AV=10
'---------------文件结束-------------
' 其中每个参数的意义如前ConfigType所定义
Dim line_str As String
Dim point1 As Integer
Dim ParName As String
Dim value_str As String
Dim value_num  As Double
Dim SetParNumber As Integer
' 按照默认设置配置MyConfig
With MyConfig
    .n_all = 20
    .n_live = 4
    .P_AV = 10
    .P_BY = 10
    .par_num = 3
    .turns = 20
End With
' 清屏
' Debug不能在代码中被清屏
Debug.Clear

' 开始读取并处理配置文件
Open filename For Input As #1
SetParNumber = 0
Do While Not EOF(1)
    On Error Resume Next
    Line Input #1, line_str
    point1 = InStr(line_str, "=")
    ParName = Mid(line_str, 1, point1 - 1)
    value_str = Mid(line_str, point1 + 1, Len(line_str) - point1)
    value_num = Val(value_str)
    ' 按照配置文件对遗传算法的各参数进行配置
    Select Case ParName
    Case "n_all"
        MyConfig.n_all = Int(value_num)
        SetParNumber = SetParNumber + 1
    Case "n_live"
        MyConfig.n_live = Int(value_num)
        SetParNumber = SetParNumber + 1
    Case "P_AV"
        MyConfig.P_AV = value_num
        SetParNumber = SetParNumber + 1
    Case "P_BY"
        MyConfig.P_BY = value_num
        SetParNumber = SetParNumber + 1
    Case "par_num"
        MyConfig.par_num = Int(value_num)
        SetParNumber = SetParNumber + 1
    Case "turns"
        MyConfig.turns = Int(value_num)
        SetParNumber = SetParNumber + 1
    End Select
Loop
Close #1
' 完成配置文件的读取和处理
If SetParNumber = ConfigParNumber Then ' 配置文件参数检查
    Debug.Print "通知：算法配置文件读取正确"
Else
    Debug.Print "警告：至少有一条算法配置未被设置"
End If
End Sub

Public Sub ReadParConfig(filename As String)
' 此程序用于读取参数的配置信息，仅有唯一参数FileName指示参数配置文件名，参数配置文件应当具有以下形式：
' {min}ParName1=10  参数可能的最小值
' {max}ParName1=15  参数可能的最大值
' {step}ParName1=0.5  参数求解精度
' -------------------------------
Dim SetParNumber As Integer
Dim point_left As Integer, point_right As Integer, point_equ As Integer
Dim line_str As String
Dim value_str As String
Dim value_num As Double
Dim ParName As String
Dim ParKind As String
Dim i As Integer
' 初始化参数表
For i = 1 To ParNumberMax
    ParList(i).MaxNum = -1000
    ParList(i).MinNum = -1000
    ParList(i).NameStr = "###"
    ParList(i).StepNum = -1000
Next i

' 开始读取并处理参数配置
Open filename For Input As #1
Do While Not EOF(1)
    On Error Resume Next
    Line Input #1, line_str
    point_left = InStr(line_str, "{")
    point_right = InStr(line_str, "}")
    point_equ = InStr(line_str, "=")
    ParName = Mid(line_str, point_right + 1, point_equ - point_right - 1)
    ParKind = Mid(line_str, point_left + 1, point_right - point_left - 1)
    value_str = Mid(line_str, point_equ + 1, Len(line_str) - point_equ)
    value_num = Val(value_str)

    For i = 1 To ParNumberMax
        If (ParList(i).NameStr = ParName) Or (ParList(i).NameStr = "###") Then
            ParList(i).NameStr = ParName
            Select Case ParKind
            Case "max"
                ParList(i).MaxNum = value_num
            Case "min"
                ParList(i).MinNum = value_num
            Case "step"
                ParList(i).StepNum = value_num
            End Select
            Exit For
        End If
    Next i

Loop
Close #1
' 完成参数配置读取和处理

' 输出参数配置处理信息
For i = 1 To ParNumberMax
    If ParList(i).NameStr = "###" Then
        Exit For
    End If
    If (ParList(i).MaxNum = -1000) Or (ParList(i).MinNum = -1000) Or (ParList(i).StepNum = -1000) Then
        Debug.Print "错误：参数" & ParList(i).NameStr & "配置不正确"
    End If
Next i
SetParNumber = i - 1
If SetParNumber <> MyConfig.par_num Then
    Debug.Print "错误：参数数量不匹配"
Else
    Debug.Print "通知：待优化参数信息配置正确"
End If
' 完成参数配置信息输出
End Sub

Public Sub DoGA()
Dim i As Long
' 此过程实现遗传算法
BeginForPar ' 生成初始解
' 对个体估价
For TurnAlready = 1 To MyConfig.turns
    ' 对个体估价
    For i = 1 To MyConfig.n_all
                BodyList(i).GoodValue = CalcValue(BodyList(i))
                Debug.Print "第" & i & "号个体评估完毕，具有适应度" & BodyList(i).GoodValue
                Debug.Print BodyList(i).ParValue(1) & "   " & BodyList(i).ParValue(2) & "    " & BodyList(i).ParValue(3) & "	" & BodyList(i).ParValue(4)
    Next i
    MakeNewBody ' 产生新个体
    OutputProfile ' 输出结果概况
Next TurnAlready
End Sub

Public Sub BeginForPar()
' 此过程用于对参数进行初始化产生初代个体
Dim i As Integer, j As Integer
Dim TimesMin As Long, TimesMax As Long
Dim RandomTimes As Long
Dim RandomRes As Double

Randomize
For i = 1 To MyConfig.n_all
    For j = 1 To MyConfig.par_num
        TimesMin = Int(ParList(j).MinNum / ParList(j).StepNum)
        TimesMax = Int(ParList(j).MaxNum / ParList(j).StepNum)
        RandomTimes = Int(Rnd * (TimesMax - TimesMin + 1) + TimesMin)
        RandomRes = RandomTimes * ParList(j).StepNum
        BodyList(i).ParValue(j) = RandomRes
    Next j
Next i
'-----------强制加入个体-----------------
'----------------------------------------
Debug.Print "通知：成功生成初代个体" & MyConfig.n_all & "个"
End Sub

Public Function CalcValue(body As BodyType) As Double
Dim Fre As Double
Dim HowGood As Double
Dim FreValue As Double
Dim HowGoodPhase As Double
' 此函数用于对个体适应度进行评估
'----------------------------
' 对求解进行初始化
DeleteResults
ResetAll
'-----------------------------
SetPar body ' 设定参数
MakeModel ' 建立模型
CalcAndGetResult1D body   ' 运行计算并导出1D结果
Fre = GetFre  ' 获得谐振频率
FreValue = GetFreValue ' 获取谐振频率的S参数值
If ((FreValue <= -9) And (Get1DResValue(Fre) <= -5)) Then
        '--------------------------
        ' 重新初始化求解器
        DeleteResults
        ResetAll
        '--------------------------
        SetFre Fre ' 设定求解器频率
        MakeModel ' 建立模型
        CalcAndGetResult2D ' 运行计算并导出2D/3D结果
        HowGoodPhase = PGPhase(body)  ' 返回相位适应度计算结果

        If HowGoodPhase > 2 Then
			CalcValue=PGFar(body)*100000+2
		Else
			CalcValue=HowGoodPhase
        End If
Else
        CalcValue = Abs(FreValue + Get1DResValue(Fre)) / 100
End If
End Function


Public Sub MakeNewBody()
' 此过程用于产生下一代个体，无输入参数
Dim temp As BodyType
Dim i As Long, j As Long
Dim TimesMin As Long, TimesMax As Long
Dim RandomTimes As Long
Dim RandomRes As Double
Dim RanBY As Double, RanAV As Double
Dim FatherTimes As Long, MotherTimes As Long
Dim RamFather As Long, RamMother As Long
Dim AVTimes As Long
Dim AVRes As Double
Dim RamMomOrPa As Double
Randomize
' 开始对个体按照适应度降序排列
For i = 1 To MyConfig.n_all
    For j = i To MyConfig.n_all
        If BodyList(i).GoodValue < BodyList(j).GoodValue Then
            temp = BodyList(i)
            BodyList(i) = BodyList(j)
            BodyList(j) = temp
        End If
    Next j
Next i
' 完成个体适应度降序排列

' 构建存活个体列表
For i = 1 To MyConfig.n_live
    BodyLiveList(i) = BodyList(i)
Next i
' 完成存活个体列表构建

' 开始交配产生下一代
For i = 1 To MyConfig.n_all
    ' 首先确定父本母本
    RamFather = Int(Rnd * (MyConfig.n_live - 1 + 1) + 1)
    RamMother = Int(Rnd * (MyConfig.n_live - 1 + 1) + 1)
    ' 对每个基因遗传
    For j = 1 To MyConfig.par_num
        RanBY = 100 * Rnd
        If RanBY <= MyConfig.P_BY Then ' 满足变异条件，该位点发生变异
            TimesMin = Int(ParList(j).MinNum / ParList(j).StepNum)
            TimesMax = Int(ParList(j).MaxNum / ParList(j).StepNum)
            RandomTimes = Int(Rnd * (TimesMax - TimesMin + 1) + TimesMin)
            RandomRes = RandomTimes * ParList(j).StepNum
            BodyList(i).ParValue(j) = RandomRes
        Else  ' 不满足变异条件
            RanAV = 100 * Rnd
            If RanAV < MyConfig.P_AV Then ' 满足平均化遗传条件，该位点平均化
                FatherTimes = Int(BodyLiveList(RamFather).ParValue(j) / ParList(j).StepNum)
                MotherTimes = Int(BodyLiveList(RamMother).ParValue(j) / ParList(j).StepNum)
                AVTimes = Int((FatherTimes + MotherTimes) / 2)
                AVRes = AVTimes * ParList(j).StepNum
                BodyList(i).ParValue(j) = AVRes
            Else  ' 既不平均化也不变异，随机从双亲继承特性
            	RamMomOrPa=100*Rnd
            	If RamMomOrPa>=50 Then
					BodyList(i).ParValue(j) = BodyLiveList(RamMother).ParValue(j)
				Else
					BodyList(i).ParValue(j) = BodyLiveList(RamFather).ParValue(j)
            	End If

            End If
        End If
    Next j
Next i
BodyList(1) = BodyLiveList(1)  ' 确保上一代最优解的存活
BestBodyList(TurnAlready) = BodyLiveList(1)
' 完成下一代产生
End Sub

Public Sub OutputProfile()
Dim i As Long
Debug.Print "---------------------------------"
Debug.Print "通知：当前完成第" & TurnAlready & "代计算，最优结果为" & BodyLiveList(1).GoodValue
Debug.Print "该个体各参数如下："
For i = 1 To MyConfig.par_num
    Debug.Print ParList(i).NameStr & " = " & BodyLiveList(1).ParValue(i)
Next i
End Sub
Public Function PGPhase(body As BodyType) As Double
' 按照相位对适应度进行评估
        Dim r As Double
        Dim fai As Double
        Dim total As Double
        Dim avg As Double
        Dim h As Double
        Dim i As Long
        Dim DataRes() As String
        Dim IndexNum As Long
        Dim DataLine() As Double
        Dim RecordNum As Long
        Dim x As Double, y As Double, z As Double
        Dim PhaseZ(1 To 360) As Double
        Dim PhaseD(1 To 359) As Double
        Dim MaxValue As Double
		Dim zb As Boolean
        ReadDataFile TempDir & "temp2D.sig", DataRes, RecordNum
        zb=False
        h = body.ParValue(5)/2
        r = body.ParValue(3)+body.ParValue(2)+body.ParValue(1)-0.5
        For i = 1 To 360
                fai = i / 360 * 2 * 3.14159
                x = 45 + r * Cos(fai)
                y = 45 + r * Sin(fai)
                z = h
                IndexNum = GetNearestPoint(DataRes, x, y, z)
                GetDataValue DataRes(IndexNum), DataLine
                PhaseZ(i) = Atn(DataLine(6) /( DataLine(9)+0.00001))
        Next i
        For i = 1 To 358
            PhaseZ(i) = (PhaseZ(i) + PhaseZ(i + 1) + PhaseZ(i + 2)) / 3
        Next i
        For i = 1 To 359
            PhaseD(i) = PhaseZ(i + 1) - PhaseZ(i)
        Next i
        total = 0
        MaxValue = 0
        Dim BigNum As Long
        BigNum = 0
        For i = 1 To 359
            If PhaseD(i) > 0 Then
                BigNum = BigNum + 1
            End If
        Next i

        If BigNum >= 50 Then ' 超过50个正数差值表明是驻波
            zb=True
        End If

        For i = 1 To 359
                If Abs(PhaseD(i)) > MaxValue Then
                        MaxValue = Abs(PhaseD(i))
                End If
        Next i

        For i = 1 To 359
                If Abs(PhaseD(i)) <= 2 Then
                    total = total + PhaseD(i) / MaxValue
                End If

        Next i

        avg = total / 359
        total = 0
        For i = 1 To 359
                total = (PhaseD(i) - avg) ^ 2 + total
        Next i
        avg = total / 359
        avg = Sqr(avg)
        If zb=False Then
			PGPhase = 1 / (avg + 0.00001)
		Else
			PGPhase = (1 / (avg + 0.00001))/5
        End If

End Function
Public Function PGFar(body As BodyType) As Double
' 按照远场对适应度进行评估

        Dim fai As Double
        Dim MaxGain As Double
        Dim MaxGainPlace As Long
        Dim h As Double
        Dim i As Long
        Dim DataRes() As String
        Dim IndexNum As Long
        Dim DataLine() As Double
        Dim RecordNum As Long
        Dim width As Double
        Dim place As Double
        Dim Gain(0 To 360) As Double
        Dim MaxValue As Double
		Dim H_right As Integer
		Dim H_left As Integer
		Dim theta As Integer

        ReadDataFile TempDir & "tempFar.sig", DataRes, RecordNum
        fai = 0

        For i = 1 To RecordNum
                GetDataValue DataRes(i), DataLine
                If Abs(DataLine(2)-fai)<=0.05 Then
					theta=Int(DataLine(1))
					Gain(theta)=DataLine(3)
                End If
        Next i
		MaxGain=-10000
		For i=1 To 360
			If Gain(i)>MaxGain Then
				MaxGain=Gain(i)
				MaxGainPlace=i
			End If
		Next i

		For i=MaxGainPlace To 360
			If Gain(i)<(MaxGain-3) Then
				Exit For
			End If
		Next i
		H_right=i

		For i=MaxGainPlace To 1 STEP -1
			If Gain(i)<(MaxGain-3) Then
				Exit For
			End If
		Next i
		H_left=i

		width=Abs(H_right-H_left)
		place=MaxGainPlace Mod 180

		PGFar=1/((Abs(place-90)+0.0001)*width)
End Function
Public Function PG(body As BodyType) As Double
' 按照强度对适应度进行评估
        Dim r As Double
        Dim fai As Double
        Dim total As Double
        Dim avg As Double
        Dim h As Double
        Dim i As Long
        Dim DataRes() As String
        Dim IndexNum As Long
        Dim DataLine() As Double
        Dim RecordNum As Long
        Dim x As Double, y As Double, z As Double
        Dim QDZ(1 To 360) As Double
        Dim MaxValue As Double

        ReadDataFile TempDir & "temp2D.sig", DataRes, RecordNum
        h = 1.575 / 2
        r = 17
        For i = 1 To 360
                fai = i / 360 * 2 * 3.14159
                x = 45 + r * Cos(fai)
                y = 45 + r * Sin(fai)
                z = h
                IndexNum = GetNearestPoint(DataRes, x, y, z)
                GetDataValue DataRes(IndexNum), DataLine
                QDZ(i) = Sqr(DataLine(6) ^ 2 + DataLine(9) ^ 2)
        Next i
        total = 0
        MaxValue = 0
        For i = 1 To 360
                If QDZ(i) > MaxValue Then
                        MaxValue = QDZ(i)
                End If
        Next i
        For i = 1 To 360
                QDZ(i) = QDZ(i) / MaxValue
        Next i
        For i = 1 To 360
                total = total + QDZ(i)
        Next i
        avg = total / 360
        total = 0
        For i = 1 To 360
                total = (QDZ(i) - avg) ^ 2 + total
        Next i
        avg = total / 360
        avg = Sqr(avg)
        PG = 1 / (avg + 0.00001)
End Function
Public Sub Export1DResult(ResultName As String, filename As String)
' 此程序用于导出1D结果，
' 参数ResultName为结果树中的完整路径，其形如“1D Results\S-Parameters\S1,1[1.0,0.0]+2[1.0,90]+3[1.0,180]+4[1.0,270],[10]”
' 参数FileName指明导出的文件名，如果文件已存在则将会覆盖
Dim flag As Boolean

SelectTreeItem (ResultName)
With Plot1D
        flag = .PlotView("magnitudedb")
        .Plot
End With
With ASCIIExport
     .FileName (TempDir & filename)
     .Execute
End With
End Sub
Public Sub Export2DResult(ResultName As String, filename As String)
' 此程序用于导出2D结果，
' 参数ResultName为结果树中的完整路径，其形如“2D/3D Results\E-Field\e-field (f=f) [1[1.0,0.0]+2[1.0,90],[10]]”
' 参数FileName指明导出的文件名，如果文件已存在则将会覆盖

        SelectTreeItem (ResultName)
        Plot3DPlotsOn2DPlane (True)
' Plot the field of the selected monitor
        With VectorPlot2D
                .Color (True)
        .Arrows (400)
        .ArrowSize (50)
        .PlaneNormal ("z")
        .PlaneCoordinate (0.78)
        .LogScale (False)
        .Plot
        End With
        With ASCIIExport
        .Reset
        .FileName (TempDir & filename)
        .Mode ("FixedWidth")
        .StepX (0.5)
        .StepY (0.5)
        .StepZ (0.8)
        .Execute
        End With

End Sub
Public Sub SetPar(body As BodyType)
' 此过程对参数进行设置，唯一参数Body表示待评估个体
Dim i As Long
For i = 1 To MyConfig.par_num
    StoreDoubleParameter(ParList(i).NameStr, body.ParValue(i))
Next i
End Sub
Public Sub CalcAndGetResult1D(body As BodyType)
' 此过程用于运行仿真并获取1D结果
On Error GoTo ErrDeal
Solver.Start
Export1DResult Mon1D, "temp1D.sig"
Export1DResult Mon1Dp, "temp1Dp.sig"
Exit Sub

ErrDeal:
' 对求解进行初始化
DeleteResults
ResetAll
'-----------------------------
Debug.Print "通知：遇到一次错误，但已修复"
MakeModel ' 建立模型
Solver.Start
Export1DResult Mon1D, "temp1D.sig"
Export1DResult Mon1Dp, "temp1Dp.sig"
End Sub
Public Sub CalcAndGetResult2D()
' 此过程用于运行仿真并获取2D结果
On Error GoTo ErrDeal
Solver.Start
Export2DResult Mon2D, "temp2D.sig"
export_farfield TempDir & "tempFar.sig"
Exit Sub

ErrDeal:
        Debug.Print "通知：遇到一次错误，但已修复"
        DeleteResults
        ResetAll
        '--------------------------
        MakeModel ' 建立模型
        Solver.Start
        Export2DResult Mon2D, "temp2D.sig"
        export_farfield TempDir & "tempFar.sig"
End Sub
Public Function GetFreP() As Double
' 此函数用于计算1Dp结果的谐振频率
Dim n As Long, i As Long
Dim vx As Double, vy As Double
Dim min As Double, f As Double

min = 100
With Result1D(TempDir & "temp1Dp.sig")
     n = .GetN              ' Get number of frequency samples
     For i = 0 To n - 1
            vx = .GetX(i)   ' Get frequency of data point
            vy = .GetY(i)   ' Get phase of data point
            If vy < min Then
                min = vy
                f = vx
            End If
     Next i
End With
GetFreP = f
End Function

Public Function GetFre() As Double
' 此函数用于计算1D结果的谐振频率
Dim n As Long, i As Long
Dim vx As Double, vy As Double
Dim min As Double, f As Double

min = 100
With Result1D(TempDir & "temp1D.sig")
     n = .GetN              ' Get number of frequency samples
     For i = 0 To n - 1
            vx = .GetX(i)   ' Get frequency of data point
            vy = .GetY(i)   ' Get phase of data point
            If vy < min Then
                min = vy
                f = vx
            End If
     Next i
End With
GetFre = f
End Function
Public Function GetFreValue() As Double
' 此函数用于计算1D结果的谐振频率最小点的S参数值
Dim n As Long, i As Long
Dim vx As Double, vy As Double
Dim min As Double, f As Double

min = 100
With Result1D(TempDir & "temp1D.sig")
     n = .GetN              ' Get number of frequency samples
     For i = 0 To n - 1
            vx = .GetX(i)   ' Get frequency of data point
            vy = .GetY(i)   ' Get phase of data point
            If vy < min Then
                min = vy
                f = vx
            End If
     Next i
End With
GetFreValue = min
End Function
Public Function Get1DResValue(Fre As Double) As Double
' 此函数用于获取指定频率处另一个1D参数图的值
Dim n As Long, i As Long
Dim vx As Double, vy As Double
Dim min As Double, f As Double

min = 100
With Result1D(TempDir & "temp1Dp.sig")
     n = .GetN              ' Get number of frequency samples
     For i = 0 To n - 1
            vx = .GetX(i)   ' Get frequency of data point
            vy = .GetY(i)   ' Get phase of data point
            If Abs(vx - Fre) <= 0.01 Then
                Exit For
            End If
     Next i
End With
Get1DResValue = vy
End Function

Public Sub export_farfield(filename As String)
' 此过程用于导出远场结果，唯一参数filename指定保存文件的完整路径
	'===========================变量定义==================================================
	Dim cst_tree() As String, cst_tmpstr As String
	Dim cst_iloop As Long, cst_iloop2 As Long
	Dim cst_nff As Long, cst_nom As Long
	Dim cst_ffq As String

	Dim cst_theta_start As Double, cst_theta_step As Double, cst_theta_stop As Double, cst_ntheta As Long
	Dim cst_phi_start As Double, cst_phi_step As Double, cst_phi_stop As Double, cst_nphi As Long
	Dim cst_phi As Double, cst_theta As Double, cst_theta_calc As Double, cst_phi_calc As Double

	Dim cst_dt As Double, cst_dp As Double

	Dim cst_icomp As Integer
	Dim cst_floop_start As Long, cst_floop_end As Long, cst_ifloop As Long

	Dim cst_title As String, cst_title_ini As String, cst_ffname As String, cst_filename As String
	Dim CST_FN As Long, cst_iff As Long
	Dim cst_Abs_Gain As Double

	Dim cst_ffpol As String

	Dim cst_xval As Double, cst_yval As Double, cst_zval As Double
	Dim cst_time As Double
	'==================================================================================
	' 遍历获取所有远场结果名称并存到cst_tree字符串数组中
	ReDim cst_tree(1)
	cst_tree(0) = Resulttree.GetFirstChildName ("Farfields")
	cst_iloop = 0
	Do
		cst_tmpstr = Resulttree.GetNextItemName(cst_tree(cst_iloop))
		If cst_tmpstr <> "" Then
			cst_iloop = cst_iloop+1
			ReDim Preserve cst_tree(cst_iloop)
			cst_tree(cst_iloop)=cst_tmpstr
		End If
	Loop Until cst_tmpstr = ""
	' 远场结果遍历完成
	' 开始去除远场结果名称中的"farfiled"字段
	cst_nff = cst_iloop+1

	For cst_iloop = 1 To cst_nff
		cst_tmpstr = Replace(cst_tree(cst_iloop-1),"Farfields\","")
		cst_tree(cst_iloop-1)=cst_tmpstr
	Next cst_iloop
	' 远场结果名称字符串处理完成

	' 获取监视器频率置于cst_freq中
	ReDim cst_freq(cst_nff)
	cst_nom = Monitor.GetNumberOfMonitors
	For cst_iloop = 1 To cst_nom
		If Monitor.GetMonitorTypeFromIndex(cst_iloop-1) = "Farfield" Then
			For cst_iloop2 = 1 To cst_nff
				If InStr(cst_tree(cst_iloop2-1),Monitor.GetMonitorNameFromIndex(cst_iloop-1)) > 0 Then
					cst_freq(cst_iloop2-1) = Monitor.GetMonitorFrequencyFromIndex(cst_iloop-1)
				End If
			Next cst_iloop2
		End If
	Next cst_iloop
	' 获取监视器频率完毕

	cst_theta_start = 0
	cst_theta_step  = 1
	cst_theta_stop = 360
	cst_ntheta = IIf(cst_theta_step=0,0,Abs(cst_theta_stop-cst_theta_start)/cst_theta_step)
	cst_phi_start = 0
	cst_phi_step = 180
	cst_phi_stop = 360
	cst_nphi = IIf(cst_phi_step=0,0,Abs(cst_phi_stop-cst_phi_start)/cst_phi_step)




	'--- check theta and phi settings
	cst_dt = cst_theta_stop - cst_theta_start
	cst_dp = cst_phi_stop - cst_phi_start
	cst_icomp  = 1
	'--- set farfield evaluation points
	With FarfieldPlot
		.Reset
		.SetPlotMode "Gain"
		.SetScaleLinear "False" 'Export performed in dB
		.SetAutomaticCoordinateSystem "True"
		.SetCoordinateSystemType "spherical"
		.SetAnntenaType "unknown"
	End With

'------------write result list-------------
	For cst_theta = cst_theta_start To cst_theta_stop STEP cst_theta_step
		For cst_phi = cst_phi_start To cst_phi_stop STEP cst_phi_step
				cst_F_tpcalc (cst_theta,cst_phi,cst_theta_calc,cst_phi_calc)
				FarfieldPlot.AddListItem(cst_theta_calc, cst_phi_calc, 0)
			cst_phi = Round (cst_phi,2)
		Next cst_phi
		cst_theta = Round (cst_theta,2)
	Next cst_theta
'-------------write data to file-----------------------------
	cst_floop_start = 0
	cst_floop_end = cst_nff - 1
	For cst_ifloop = cst_floop_start To cst_floop_end

		cst_title = ""
		'--- farfield plot init
		SelectTreeItem "Farfields\"+cst_tree(cst_ifloop)
		FarfieldPlot.CalculateList(cst_tree(cst_ifloop))
		cst_ffname = cst_tree(cst_ifloop)
'		cst_filename = GetProjectPath("Result") + "CSTFarfield_WinProp_" + "(" + cst_ffname + ").apa"
		CST_FN = FreeFile
		Open filename For Output As #CST_FN
		Print #CST_FN,"* Theta        Phi        Attenuation"
		Print #CST_FN,"* Theta        Phi        Attenuation"
		cst_iff = 0
		For cst_theta = cst_theta_start To cst_theta_stop STEP cst_theta_step
			For cst_phi = cst_phi_start To cst_phi_stop STEP cst_phi_step
				'--- transversal components
				If cst_icomp = 1 Then
					'--- spherical coordinates: theta, phi
					'Dim cst_t As String, cst_p As String
					'cst_t = FarfieldPlot.GetListItem(cst_iff,"Point_T")
					'cst_p = FarfieldPlot.GetListItem(cst_iff,"Point_P")
					cst_Abs_Gain = FarfieldPlot.GetListItem(cst_iff,"abs")
				End If
				'--- write out farfield patterns for phi variation
					Print #CST_FN, " " + cstreh_Pretty_two(cst_theta) + cstreh_Pretty_two(cst_phi) + cstreh_Pretty_four(cst_Abs_Gain)
				cst_iff = cst_iff + 1
			'used Dround to overcome numerical error
			cst_phi = Round (cst_phi,2)
			Next cst_phi
			cst_theta = Round (cst_theta,2)
			Next cst_theta
		Close #CST_FN
	Next cst_ifloop
End Sub
Sub cst_F_tpcalc(ByVal theta As Double, ByVal phi As Double, ByRef tout As Double, ByRef pout As Double)

	While theta < 0
		theta = theta + 360
	Wend
	While theta > 360
		theta = theta - 360
	Wend

	If theta > 180 Then
		theta = 360 - theta
		phi = phi + 180
	End If

	While phi < 0
		phi = phi + 360
	Wend
	While phi > 360
		phi = phi - 360
	Wend

	tout = theta
	pout = phi

End Sub
Function cstreh_Pretty_two(x As Variant) As String

        'cstreh_Pretty = Replace(Left$(IIf(Left$(CStr(x), 1) = "-", "", " ") + Format(x,"0.0000000000E+00") + String$(16, " "), 18), ",", ".")
        cstreh_Pretty_two = Replace(Left$(IIf(Left$(CStr(x), 1) = "-", "", " ") + Format(x,"0.00") + String$(10, " "), 12), ",", ".")

End Function

Function cstreh_Pretty_four(x As Variant) As String

        cstreh_Pretty_four = Replace(Left$(IIf(Left$(CStr(x), 1) = "-", "", " ") + Format(x,"0.0000") + String$(10, " "), 12), ",", ".")

End Function


Public Sub SetFre(Fre As Double)
' 此过程用于设定监视器频率，唯一参数Fre表示监视器频率,必须要求监视器频率为参数f
StoreDoubleParameter("f", Fre)
End Sub

Public Sub ReadDataFile(filename As String, ByRef Datas() As String, ByRef MaxNumber As Long)
' 此过程用于将ASCII文件数据读取至Datas参数所传递的字符串数组中，MaxNumber返回了数据点的个数，filename为数据文件文件名
Dim Lines As String
Dim i As Long
Open filename For Input As #1
    Do While Not EOF(1)
        On Error Resume Next
        Line Input #1, Lines
        i = i + 1
        ReDim Preserve Datas(0 To i) As String
        Datas(i) = Lines
    Loop
Close #1
MaxNumber = i
End Sub

Public Sub GetDataValue(DataStr As String, ByRef DataValue() As Double)
' 此过程用于将文件中的行数据分割为Double型数组，DataStr为行数据，DataValue为分割结果
    Dim Max As Long
    Dim i As Long, k As Long
    Dim s() As String
    s() = Split(DataStr, " ")
    Max = UBound(s)
    k = 0
    For i = 0 To Max
      If s(i) <> "" Then
        k = k + 1
        ReDim Preserve DataValue(1 To k) As Double
        DataValue(k) = Val(s(i))
      End If
    Next i
End Sub

Public Function GetNearestPoint(Datas() As String, x As Double, y As Double, z As Double) As Long
' 此函数用于获取文件中有记录的最靠近待求点的记录条目编号，Datas为读取的数据集，X,Y,Z为待求点的坐标
Dim xp As Double, yp As Double, zp As Double
Dim DisMin As Double
Dim Dis As Double
Dim ResIndex As Long
Dim i As Long
Dim MaxNum As Long
Dim nums() As Double

DisMin = 100000
MaxNum = UBound(Datas)
For i = 3 To MaxNum
    GetDataValue Datas(i), nums
    xp = nums(1)
    yp = nums(2)
    zp = nums(3)
    Dis = Sqr((x - xp) ^ 2 + (y - yp) ^ 2 + (z - zp) ^ 2)
    If Dis < DisMin Then
        ResIndex = i
        DisMin = Dis
    End If
Next i
GetNearestPoint = ResIndex
End Function


Public Sub MakeModel()
' 此过程完成全部的建模过程


End Sub
