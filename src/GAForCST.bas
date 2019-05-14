' GNU GENERAL PUBLIC LICENSE
'                      Version 3, 29 June 2007
' Copyright (C) 2007 Free Software Foundation, Inc. <https://fsf.org/>
' Everyone is permitted to copy and distribute verbatim copies
' of this license document, but changing it is not allowed.

Option Explicit
' ��������
Public Const ConfigParNumber = 6 ' ��Ҫ���õ��Ŵ��㷨��������
Public Const ParNumberMax = 100 ' �������Ĵ��Ż�������Ŀ
Public Const TurnNumberMax = 1000 ' �Ŵ����������ֵ
Public Const TempDir = ""  ' ��ʱ�ļ����Ŀ¼�����磺E:\qtemp\
Public Const Mon1D = "1D Results\S-Parameters\S1,1[1.0,0.0]+2[1.0,90],[10]" ' 1D����������
Public Const Mon1Dp = "1D Results\S-Parameters\S2,1[1.0,0.0]+2[1.0,90],[10]" ' ��һ��1D����������
Public Const Mon2D = "2D/3D Results\E-Field\e-field (f=f) [1[1.0,0.0]+2[1.0,90],[10]]" ' 2D����������
' ��ɳ�������

' ����ConfigType���Ͷ����ý��д洢
Type ConfigType
    par_num As Long ' ���Ż���������
    n_live As Long ' ÿһ�����ĸ�������
    n_all As Long ' ÿһ���ĸ�������
    turns As Long ' ������
    P_BY As Double ' ��������ĸ���
    P_AV As Double ' ����ƽ�����Ŵ��ĸ���
End Type
' ���ConfigType���͵Ķ���

' ����ParType���ͶԴ��Ż��������д洢
Type ParType
    NameStr As String ' ������
    MinNum As Double ' ������Сֵ
    MaxNum As Double ' �������ֵ
    StepNum As Double ' ��������
End Type
' ���ParType���͵Ķ���

' ����BodyType���ͶԸ�����д洢
Type BodyType
    ParValue(1 To ParNumberMax) As Double ' ����ParList˳��洢�ĸ������ֵ
    GoodValue As Double ' ����õ�����Ӧ��
End Type
' ���BodyType���͵Ķ���



' ���幫������
Public MyConfig As ConfigType ' MyConfig�洢�����Ŵ��㷨��������Ϣ
Public ParList(1 To ParNumberMax) As ParType ' ParList�洢���Ż��Ĳ����б�
Public BodyList(1 To 1000) As BodyType ' BodyList�洢��Ⱥ�����б�
Public BodyLiveList(1 To 1000) As BodyType ' BodyLiveList�洢�������б�
Public TurnAlready As Long ' TurnAlready�洢�Ŵ��㷨������
Public BestBodyList(1 To TurnNumberMax) As BodyType ' BestBodyList�洢ÿһ�����Ž�����б�
'--------------------------------------------
Sub Main()
Dim a As Double
Dim body As BodyType


ReadConfig "E:\CYQ\�Ŵ��㷨����\config.txt"
ReadParConfig "E:\CYQ\�Ŵ��㷨����\par_config.txt"
DoGA

End Sub
'--------------------------------------------

Public Sub ReadConfig(filename As String)
' �˹������ڶ�ȡ�����ļ���Ψһ����FileNameָʾ�����ļ�·���������ļ�Ӧ�������¸�ʽ��
'------------------------------------
' par_num=3
' n_live=10
' n_all=100
' turns=50
' P_BY=10
' P_AV=10
'---------------�ļ�����-------------
' ����ÿ��������������ǰConfigType������
Dim line_str As String
Dim point1 As Integer
Dim ParName As String
Dim value_str As String
Dim value_num  As Double
Dim SetParNumber As Integer
' ����Ĭ����������MyConfig
With MyConfig
    .n_all = 20
    .n_live = 4
    .P_AV = 10
    .P_BY = 10
    .par_num = 3
    .turns = 20
End With
' ����
' Debug�����ڴ����б�����
Debug.Clear

' ��ʼ��ȡ�����������ļ�
Open filename For Input As #1
SetParNumber = 0
Do While Not EOF(1)
    On Error Resume Next
    Line Input #1, line_str
    point1 = InStr(line_str, "=")
    ParName = Mid(line_str, 1, point1 - 1)
    value_str = Mid(line_str, point1 + 1, Len(line_str) - point1)
    value_num = Val(value_str)
    ' ���������ļ����Ŵ��㷨�ĸ�������������
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
' ��������ļ��Ķ�ȡ�ʹ���
If SetParNumber = ConfigParNumber Then ' �����ļ��������
    Debug.Print "֪ͨ���㷨�����ļ���ȡ��ȷ"
Else
    Debug.Print "���棺������һ���㷨����δ������"
End If
End Sub

Public Sub ReadParConfig(filename As String)
' �˳������ڶ�ȡ������������Ϣ������Ψһ����FileNameָʾ���������ļ��������������ļ�Ӧ������������ʽ��
' {min}ParName1=10  �������ܵ���Сֵ
' {max}ParName1=15  �������ܵ����ֵ
' {step}ParName1=0.5  ������⾫��
' -------------------------------
Dim SetParNumber As Integer
Dim point_left As Integer, point_right As Integer, point_equ As Integer
Dim line_str As String
Dim value_str As String
Dim value_num As Double
Dim ParName As String
Dim ParKind As String
Dim i As Integer
' ��ʼ��������
For i = 1 To ParNumberMax
    ParList(i).MaxNum = -1000
    ParList(i).MinNum = -1000
    ParList(i).NameStr = "###"
    ParList(i).StepNum = -1000
Next i

' ��ʼ��ȡ�������������
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
' ��ɲ������ö�ȡ�ʹ���

' ����������ô�����Ϣ
For i = 1 To ParNumberMax
    If ParList(i).NameStr = "###" Then
        Exit For
    End If
    If (ParList(i).MaxNum = -1000) Or (ParList(i).MinNum = -1000) Or (ParList(i).StepNum = -1000) Then
        Debug.Print "���󣺲���" & ParList(i).NameStr & "���ò���ȷ"
    End If
Next i
SetParNumber = i - 1
If SetParNumber <> MyConfig.par_num Then
    Debug.Print "���󣺲���������ƥ��"
Else
    Debug.Print "֪ͨ�����Ż�������Ϣ������ȷ"
End If
' ��ɲ���������Ϣ���
End Sub

Public Sub DoGA()
Dim i As Long
' �˹���ʵ���Ŵ��㷨
BeginForPar ' ���ɳ�ʼ��
' �Ը������
For TurnAlready = 1 To MyConfig.turns
    ' �Ը������
    For i = 1 To MyConfig.n_all
                BodyList(i).GoodValue = CalcValue(BodyList(i))
                Debug.Print "��" & i & "�Ÿ���������ϣ�������Ӧ��" & BodyList(i).GoodValue
                Debug.Print BodyList(i).ParValue(1) & "   " & BodyList(i).ParValue(2) & "    " & BodyList(i).ParValue(3) & "	" & BodyList(i).ParValue(4)
    Next i
    MakeNewBody ' �����¸���
    OutputProfile ' �������ſ�
Next TurnAlready
End Sub

Public Sub BeginForPar()
' �˹������ڶԲ������г�ʼ��������������
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
'-----------ǿ�Ƽ������-----------------
'----------------------------------------
Debug.Print "֪ͨ���ɹ����ɳ�������" & MyConfig.n_all & "��"
End Sub

Public Function CalcValue(body As BodyType) As Double
Dim Fre As Double
Dim HowGood As Double
Dim FreValue As Double
Dim HowGoodPhase As Double
' �˺������ڶԸ�����Ӧ�Ƚ�������
'----------------------------
' �������г�ʼ��
DeleteResults
ResetAll
'-----------------------------
SetPar body ' �趨����
MakeModel ' ����ģ��
CalcAndGetResult1D body   ' ���м��㲢����1D���
Fre = GetFre  ' ���г��Ƶ��
FreValue = GetFreValue ' ��ȡг��Ƶ�ʵ�S����ֵ
If ((FreValue <= -9) And (Get1DResValue(Fre) <= -5)) Then
        '--------------------------
        ' ���³�ʼ�������
        DeleteResults
        ResetAll
        '--------------------------
        SetFre Fre ' �趨�����Ƶ��
        MakeModel ' ����ģ��
        CalcAndGetResult2D ' ���м��㲢����2D/3D���
        HowGoodPhase = PGPhase(body)  ' ������λ��Ӧ�ȼ�����

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
' �˹������ڲ�����һ�����壬���������
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
' ��ʼ�Ը��尴����Ӧ�Ƚ�������
For i = 1 To MyConfig.n_all
    For j = i To MyConfig.n_all
        If BodyList(i).GoodValue < BodyList(j).GoodValue Then
            temp = BodyList(i)
            BodyList(i) = BodyList(j)
            BodyList(j) = temp
        End If
    Next j
Next i
' ��ɸ�����Ӧ�Ƚ�������

' �����������б�
For i = 1 To MyConfig.n_live
    BodyLiveList(i) = BodyList(i)
Next i
' ��ɴ������б���

' ��ʼ���������һ��
For i = 1 To MyConfig.n_all
    ' ����ȷ������ĸ��
    RamFather = Int(Rnd * (MyConfig.n_live - 1 + 1) + 1)
    RamMother = Int(Rnd * (MyConfig.n_live - 1 + 1) + 1)
    ' ��ÿ�������Ŵ�
    For j = 1 To MyConfig.par_num
        RanBY = 100 * Rnd
        If RanBY <= MyConfig.P_BY Then ' ���������������λ�㷢������
            TimesMin = Int(ParList(j).MinNum / ParList(j).StepNum)
            TimesMax = Int(ParList(j).MaxNum / ParList(j).StepNum)
            RandomTimes = Int(Rnd * (TimesMax - TimesMin + 1) + TimesMin)
            RandomRes = RandomTimes * ParList(j).StepNum
            BodyList(i).ParValue(j) = RandomRes
        Else  ' �������������
            RanAV = 100 * Rnd
            If RanAV < MyConfig.P_AV Then ' ����ƽ�����Ŵ���������λ��ƽ����
                FatherTimes = Int(BodyLiveList(RamFather).ParValue(j) / ParList(j).StepNum)
                MotherTimes = Int(BodyLiveList(RamMother).ParValue(j) / ParList(j).StepNum)
                AVTimes = Int((FatherTimes + MotherTimes) / 2)
                AVRes = AVTimes * ParList(j).StepNum
                BodyList(i).ParValue(j) = AVRes
            Else  ' �Ȳ�ƽ����Ҳ�����죬�����˫�׼̳�����
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
BodyList(1) = BodyLiveList(1)  ' ȷ����һ�����Ž�Ĵ��
BestBodyList(TurnAlready) = BodyLiveList(1)
' �����һ������
End Sub

Public Sub OutputProfile()
Dim i As Long
Debug.Print "---------------------------------"
Debug.Print "֪ͨ����ǰ��ɵ�" & TurnAlready & "�����㣬���Ž��Ϊ" & BodyLiveList(1).GoodValue
Debug.Print "�ø�����������£�"
For i = 1 To MyConfig.par_num
    Debug.Print ParList(i).NameStr & " = " & BodyLiveList(1).ParValue(i)
Next i
End Sub
Public Function PGPhase(body As BodyType) As Double
' ������λ����Ӧ�Ƚ�������
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

        If BigNum >= 50 Then ' ����50��������ֵ������פ��
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
' ����Զ������Ӧ�Ƚ�������

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
' ����ǿ�ȶ���Ӧ�Ƚ�������
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
' �˳������ڵ���1D�����
' ����ResultNameΪ������е�����·���������硰1D Results\S-Parameters\S1,1[1.0,0.0]+2[1.0,90]+3[1.0,180]+4[1.0,270],[10]��
' ����FileNameָ���������ļ���������ļ��Ѵ����򽫻Ḳ��
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
' �˳������ڵ���2D�����
' ����ResultNameΪ������е�����·���������硰2D/3D Results\E-Field\e-field (f=f) [1[1.0,0.0]+2[1.0,90],[10]]��
' ����FileNameָ���������ļ���������ļ��Ѵ����򽫻Ḳ��

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
' �˹��̶Բ����������ã�Ψһ����Body��ʾ����������
Dim i As Long
For i = 1 To MyConfig.par_num
    StoreDoubleParameter(ParList(i).NameStr, body.ParValue(i))
Next i
End Sub
Public Sub CalcAndGetResult1D(body As BodyType)
' �˹����������з��沢��ȡ1D���
On Error GoTo ErrDeal
Solver.Start
Export1DResult Mon1D, "temp1D.sig"
Export1DResult Mon1Dp, "temp1Dp.sig"
Exit Sub

ErrDeal:
' �������г�ʼ��
DeleteResults
ResetAll
'-----------------------------
Debug.Print "֪ͨ������һ�δ��󣬵����޸�"
MakeModel ' ����ģ��
Solver.Start
Export1DResult Mon1D, "temp1D.sig"
Export1DResult Mon1Dp, "temp1Dp.sig"
End Sub
Public Sub CalcAndGetResult2D()
' �˹����������з��沢��ȡ2D���
On Error GoTo ErrDeal
Solver.Start
Export2DResult Mon2D, "temp2D.sig"
export_farfield TempDir & "tempFar.sig"
Exit Sub

ErrDeal:
        Debug.Print "֪ͨ������һ�δ��󣬵����޸�"
        DeleteResults
        ResetAll
        '--------------------------
        MakeModel ' ����ģ��
        Solver.Start
        Export2DResult Mon2D, "temp2D.sig"
        export_farfield TempDir & "tempFar.sig"
End Sub
Public Function GetFreP() As Double
' �˺������ڼ���1Dp�����г��Ƶ��
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
' �˺������ڼ���1D�����г��Ƶ��
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
' �˺������ڼ���1D�����г��Ƶ����С���S����ֵ
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
' �˺������ڻ�ȡָ��Ƶ�ʴ���һ��1D����ͼ��ֵ
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
' �˹������ڵ���Զ�������Ψһ����filenameָ�������ļ�������·��
	'===========================��������==================================================
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
	' ������ȡ����Զ��������Ʋ��浽cst_tree�ַ���������
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
	' Զ������������
	' ��ʼȥ��Զ����������е�"farfiled"�ֶ�
	cst_nff = cst_iloop+1

	For cst_iloop = 1 To cst_nff
		cst_tmpstr = Replace(cst_tree(cst_iloop-1),"Farfields\","")
		cst_tree(cst_iloop-1)=cst_tmpstr
	Next cst_iloop
	' Զ����������ַ����������

	' ��ȡ������Ƶ������cst_freq��
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
	' ��ȡ������Ƶ�����

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
' �˹��������趨������Ƶ�ʣ�Ψһ����Fre��ʾ������Ƶ��,����Ҫ�������Ƶ��Ϊ����f
StoreDoubleParameter("f", Fre)
End Sub

Public Sub ReadDataFile(filename As String, ByRef Datas() As String, ByRef MaxNumber As Long)
' �˹������ڽ�ASCII�ļ����ݶ�ȡ��Datas���������ݵ��ַ��������У�MaxNumber���������ݵ�ĸ�����filenameΪ�����ļ��ļ���
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
' �˹������ڽ��ļ��е������ݷָ�ΪDouble�����飬DataStrΪ�����ݣ�DataValueΪ�ָ���
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
' �˺������ڻ�ȡ�ļ����м�¼����������ļ�¼��Ŀ��ţ�DatasΪ��ȡ�����ݼ���X,Y,ZΪ����������
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
' �˹������ȫ���Ľ�ģ����


End Sub
