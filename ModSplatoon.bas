Attribute VB_Name = "ModSplatoon"
Option Explicit

'Bloodstain                      �E�E�E���ꏊ�FVBAProject.Module1
'�����̌v�Z                      �E�E�E���ꏊ�FVBAProject.Module1
'SplineXYParaFast                �E�E�E���ꏊ�FFukamiAddins3.ModApproximate
'SplineParaFast                  �E�E�E���ꏊ�FFukamiAddins3.ModApproximate
'SplineByArrayX1DFast            �E�E�E���ꏊ�FFukamiAddins3.ModApproximate
'�X�v���C����ԍ������p�ɕ��������E�E�E���ꏊ�FFukamiAddins3.ModApproximate
'ExtractByRangeArray1D           �E�E�E���ꏊ�FFukamiAddins3.ModApproximate
'CheckArray1D                    �E�E�E���ꏊ�FFukamiAddins3.ModArray
'CheckArray1DStart1              �E�E�E���ꏊ�FFukamiAddins3.ModArray
'ExtractArray1D                  �E�E�E���ꏊ�FFukamiAddins3.ModArray
'ExtractArray                    �E�E�E���ꏊ�FFukamiAddins3.ModArray
'CheckArray2D                    �E�E�E���ꏊ�FFukamiAddins3.ModArray
'CheckArray2DStart1              �E�E�E���ꏊ�FFukamiAddins3.ModArray
'SplineByArrayX1D                �E�E�E���ꏊ�FFukamiAddins3.ModApproximate
'SplineKeisu                     �E�E�E���ꏊ�FFukamiAddins3.ModApproximate
'F_MMult                         �E�E�E���ꏊ�FFukamiAddins3.ModMatrix
'F_Minverse                      �E�E�E���ꏊ�FFukamiAddins3.ModMatrix
'�����s�񂩃`�F�b�N              �E�E�E���ꏊ�FFukamiAddins3.ModMatrix
'F_MDeterm                       �E�E�E���ꏊ�FFukamiAddins3.ModMatrix
'F_Mgyoirekae                    �E�E�E���ꏊ�FFukamiAddins3.ModMatrix
'F_Mgyohakidasi                  �E�E�E���ꏊ�FFukamiAddins3.ModMatrix
'F_Mjyokyo                       �E�E�E���ꏊ�FFukamiAddins3.ModMatrix
'UnionArray1D                    �E�E�E���ꏊ�FFukamiAddins3.ModArray
'DrawPolyLine                    �E�E�E���ꏊ�FFukamiAddins3.ModDrawShape
'GetXYDocumentFromCursor         �E�E�E���ꏊ�FFukamiAddins3.ModCursor
'GetXYCellScreenUpperLeft        �E�E�E���ꏊ�FFukamiAddins3.ModCursor
'GetPaneOfCell                   �E�E�E���ꏊ�FFukamiAddins3.ModCursor
'GetXYCellScreenLowerRight       �E�E�E���ꏊ�FFukamiAddins3.ModCursor

'------------------------------
Const Pi# = 3.141529
'------------------------------
'�V�[�g�֐��p�ߎ��A��Ԋ֐�
'------------------------------
'�z��̏����֌W�̃v���V�[�W��
'------------------------------
'�s����g�����v�Z
'��֊֐�
'------------------------------
'�V�F�C�v��}�֘A���W���[��
'20210914�쐬
'------------------------------
'������������������������������������������������������
'�J�[�\���̃X�N���[�����W�擾�p
#If VBA7 Then
Private Declare PtrSafe Function GetCursorPos Lib "user32" (IpPoint As PointAPI) As Long
#Else
Private Declare Function GetCursorPos Lib "user32" (IpPoint As PointAPI) As Long
#End If

Private Type PointAPI
    X As Long
    Y As Long
End Type
'------------------------------


Sub Bloodstain(TargetSheet As Worksheet)
'�J�[�\���ʒu�Ɍ������
'20211009

'����
'TargetSheet�E�E�E�n���΂��Ώۂ̃V�[�g/Worksheet�^

    '�J�[�\���ʒu�̃h�L�������g���W�擾
    Dim CenterX#, CenterY#
    Dim Dummy
    On Error Resume Next '�X�N���[�����W�擾�Ɏ��s�����ꍇ
    Dummy = GetXYDocumentFromCursor
    CenterX = Dummy(1)
    CenterY = Dummy(2)
    On Error GoTo 0
    If CenterX = 0 Then
        Exit Sub
    End If
    
    Dim N&, I&
    Dim k_fai1#, k_fai2#, k_fai3#, r0#, kr#, p#
    
    N = 10 + Rnd() * 6 '�����̃c�m�̌�
    k_fai1 = 0.4 '�����̃c�m�̍������̍��W�̊p�x�̌W��
    k_fai2 = 0.11 '�����̃c�m�̂��т�̍��W�̊p�x�̌W��
    k_fai3 = 0.2 '�����̃c�m�̖c��݂̍��W�̊p�x�̌W��
    r0 = 4 / 20 * N * 1.2   '�����̊j���a
    kr = 0.9 + 0.2 * Rnd() '�c�m�̒����W���B�傫���قǃc�m�������Ȃ�
    p = 0.3 '�����W��(�גʂ��̃c�m�Ƃ̊Ԋu�̃����_������)(�傫���قǊԊu���傫���ς��)
    
    '�F���X�g
    Dim ColorList&(1 To 6)
    ColorList(1) = RGB(0, 0, 255) '��
    ColorList(2) = RGB(231, 34, 231) '��
    ColorList(3) = RGB(255, 124, 0) '�I�����W
    ColorList(4) = RGB(0, 255, 255) '���F
    ColorList(5) = RGB(158, 255, 69) '����
    ColorList(6) = RGB(255, 0, 148) '��2
    
    Dim ColorNum&, InputColor&
    ColorNum = WorksheetFunction.RandBetween(1, 6)
    InputColor = ColorList(ColorNum)
    
    Call �����̌v�Z(N, k_fai1, k_fai2, k_fai3, r0, kr, p, CenterX, CenterY, TargetSheet, InputColor)

End Sub

Private Sub �����̌v�Z(N&, k_fai1#, k_fai2#, k_fai3#, r0#, kr#, p#, CenterX#, CenterY#, TargetSheet As Worksheet, Optional InputColor& = rgbRed)
'�����̌`����v�Z���āA�w��ʒu�ɕ`��
'20211009

'N          �E�E�E�����̃c�m�̌�
'k_fai1     �E�E�E�����̃c�m�̍������̍��W�̊p�x�̌W��
'k_fai2     �E�E�E�����̃c�m�̂��т�̍��W�̊p�x�̌W��
'k_fai3     �E�E�E�����̃c�m�̖c��݂̍��W�̊p�x�̌W��
'r0         �E�E�E�����̊j���a
'kr         �E�E�E�c�m�̒����W���B�傫���قǃc�m�������Ȃ�
'p          �E�E�E�����W��(�גʂ��̃c�m�Ƃ̊Ԋu�̃����_������)(�傫���قǊԊu���傫���ς��)
'CenterX    �E�E�E�����̒��SX
'CenterY    �E�E�E�����̒��SY
'TargetSheet�E�E�E�����΂��Ώۂ̃V�[�g
'InputColor �E�E�E�h��Ԃ��F�B�f�t�H���g�͐�

    Dim I&
    
    Dim ThetaList#(), ThetaDashList#()
    Dim dTheta#
    ReDim ThetaList(1 To N)
    ReDim ThetaDashList(1 To N)
    
    For I = 1 To N
        ThetaList(I) = 2 * Pi / N * I - Pi / N   '��i
        dTheta = p * Pi / N * (2 * Rnd() - 1)    'd��
        ThetaDashList(I) = ThetaList(I) + dTheta '��'i
    Next I
    
    Dim Fai1#, Fai2#, Fai3#, FaiList#(), rList#()
    Dim dr#
    Fai1 = k_fai1 * Pi / N '��1
    Fai2 = k_fai2 * Pi / N '��1
    Fai3 = k_fai3 * Pi / N '��1
    
    ReDim FaiList(1 To N, 1 To 3) '��i_1,��i_2,��i_3
    ReDim rList(1 To N, 1 To 3)   'ri_1,ri_2,ri_3
    
    For I = 1 To N
        FaiList(I, 1) = Fai1
        FaiList(I, 2) = Fai2
        FaiList(I, 3) = Fai3
        
        dr = kr * r0 * (Rnd() + 0.2)
        rList(I, 1) = r0 + dr
        rList(I, 2) = r0 + 3 / 10 * dr
        rList(I, 3) = r0 + 8 / 10 * dr
    Next I
    
    Dim XYList
    ReDim XYList(1 To 7 * N + 1, 1 To 2)
    Dim TmpIti&
    Dim TmpTheta#, Tmpr#
    For I = 1 To N
        TmpIti = (I - 1) * 7 + 1
        
        TmpTheta = ThetaDashList(I) - FaiList(I, 1)
        Tmpr = r0
        XYList(TmpIti, 1) = Tmpr * Cos(TmpTheta)
        XYList(TmpIti, 2) = Tmpr * Sin(TmpTheta)
        
        TmpTheta = ThetaDashList(I) - FaiList(I, 2)
        Tmpr = rList(I, 2)
        XYList(TmpIti + 1, 1) = Tmpr * Cos(TmpTheta)
        XYList(TmpIti + 1, 2) = Tmpr * Sin(TmpTheta)
        
        TmpTheta = ThetaDashList(I) - FaiList(I, 3)
        Tmpr = rList(I, 3)
        XYList(TmpIti + 2, 1) = Tmpr * Cos(TmpTheta)
        XYList(TmpIti + 2, 2) = Tmpr * Sin(TmpTheta)
    
        TmpTheta = ThetaDashList(I)
        Tmpr = rList(I, 1)
        XYList(TmpIti + 3, 1) = Tmpr * Cos(TmpTheta)
        XYList(TmpIti + 3, 2) = Tmpr * Sin(TmpTheta)

        TmpTheta = ThetaDashList(I) + FaiList(I, 3)
        Tmpr = rList(I, 3)
        XYList(TmpIti + 4, 1) = Tmpr * Cos(TmpTheta)
        XYList(TmpIti + 4, 2) = Tmpr * Sin(TmpTheta)

        TmpTheta = ThetaDashList(I) + FaiList(I, 2)
        Tmpr = rList(I, 2)
        XYList(TmpIti + 5, 1) = Tmpr * Cos(TmpTheta)
        XYList(TmpIti + 5, 2) = Tmpr * Sin(TmpTheta)

        TmpTheta = ThetaDashList(I) + FaiList(I, 1)
        Tmpr = r0
        XYList(TmpIti + 6, 1) = Tmpr * Cos(TmpTheta)
        XYList(TmpIti + 6, 2) = Tmpr * Sin(TmpTheta)
    
    Next I
    
    XYList(UBound(XYList, 1), 1) = XYList(1, 1)
    XYList(UBound(XYList, 1), 2) = XYList(1, 2)
    
    '�X�v���C����Ԃœ_�𑝂₷
    Dim BunkatuN&
    BunkatuN = 1000
    XYList = SplineXYParaFast(XYList, BunkatuN, 4)
    
    '��}���S�ʒu�ֈړ�
    For I = 1 To UBound(XYList, 1)
        XYList(I, 1) = CenterX + XYList(I, 1) * 10
        XYList(I, 2) = CenterY + XYList(I, 2) * 10
    Next
    
    '�o��
    Dim TmpShape As Shape
'    Application.ScreenUpdating = False
    Set TmpShape = DrawPolyLine(XYList, TargetSheet)
'    Set TmpShape = �Ȑ�����}����(XYList, CenterX, CenterY, 10)
'    Application.ScreenUpdating = True
    With TmpShape
        .Fill.ForeColor.RGB = InputColor
        .Line.ForeColor.RGB = InputColor
    End With
        
End Sub

Private Function SplineXYParaFast(ByVal ArrayXY2D, BunkatuN&, PointCount&)
'�p�����g���b�N�֐��`���ŃX�v���C����Ԃ��s��
'�������Čv�Z������������
'ArrayX,ArrayY���ǂ�����P�������A�P�������łȂ��ꍇ�ɗp����B
    
'����
'ArrayXY2D �E�E�E��Ԃ̑ΏۂƂȂ�X,Y�̒l���i�[���ꂽ�z��
'ArrayXY2D��1��ڂ�X,2��ڂ�Y�ƂȂ�悤�ɂ���B
'BunkatuN  �E�E�E�p�����g���b�N�֐��̕������i�o�͂����XList,YList�̗v�f����(������+1)�j
'PointCount�E�E�E��������ۂ̓_��
    
'�Ԃ�l
'�p�����g���b�N�֐��`���ŕ�Ԃ��ꂽXList,YList���i�[���ꂽXYList
'1��ڂ�XList,2��ڂ�YList
    
    '���͒l�̃`�F�b�N�y�яC��'������������������������������������������������������
    '���͂��Z������(���[�N�V�[�g�֐�)�������ꍇ�̏���
    If IsObject(ArrayXY2D) Then
        ArrayXY2D = ArrayXY2D.Value
    End If
        
    '�s��̊J�n�v�f��1�ɕύX�i�v�Z���₷������j
    Dim StartNum%
    StartNum = LBound(ArrayXY2D) '���͔z��̗v�f�̊J�n�ԍ�������Ă����i�o�͒l�ɍ��킹�邽�߁j
    If LBound(ArrayXY2D, 1) <> 1 Or LBound(ArrayXY2D, 2) <> 1 Then
        ArrayXY2D = Application.Transpose(Application.Transpose(ArrayXY2D))
    End If
    
    Dim ArrayX1D, ArrayY1D
    Dim I%, N%
    N = UBound(ArrayXY2D, 1)
    ReDim ArrayX1D(StartNum To StartNum - 1 + N)
    ReDim ArrayY1D(StartNum To StartNum - 1 + N)
    
    For I = 1 To N
        ArrayX1D(I + StartNum - 1) = ArrayXY2D(I, 1)
        ArrayY1D(I + StartNum - 1) = ArrayXY2D(I, 2)
    Next I
    
    '�v�Z����������������������������������������������������������
    Dim Dummy
    Dim OutputArrayX1D, OutputArrayY1D
    Dummy = SplineParaFast(ArrayX1D, ArrayY1D, BunkatuN, PointCount)
    OutputArrayX1D = Dummy(1)
    OutputArrayY1D = Dummy(2)
    
    Dim OutputArrayXY2D
    ReDim OutputArrayXY2D(StartNum To StartNum - 1 + BunkatuN + 1, 1 To 2)
    
    For I = 1 To BunkatuN + 1
        OutputArrayXY2D(StartNum + I - 1, 1) = OutputArrayX1D(StartNum + I - 1)
        OutputArrayXY2D(StartNum + I - 1, 2) = OutputArrayY1D(StartNum + I - 1)
    Next I
    
    '�o�́�����������������������������������������������������
    SplineXYParaFast = OutputArrayXY2D
    
End Function

Private Function SplineParaFast(ByVal ArrayX1D, ByVal ArrayY1D, BunkatuN&, PointCount&)
'�p�����g���b�N�֐��`���ŃX�v���C����Ԃ��s��
'�������Čv�Z������������
'ArrayX1D,ArrayY1D���ǂ�����P�������A�P�������łȂ��ꍇ�ɗp����B
'20211009

'����
'ArrayX1D  �E�E�E��Ԃ̑ΏۂƂȂ�X�̒l���i�[���ꂽ�z��
'ArrayY1D  �E�E�E��Ԃ̑ΏۂƂȂ�Y�̒l���i�[���ꂽ�z��
'BunkatuN  �E�E�E�p�����g���b�N�֐��̕������i�o�͂����OutputArrayX1D,OutputArrayY1D�̗v�f����(������+1)�j
'PointCount�E�E�E��������ۂ̓_��

'�Ԃ�l
'�p�����g���b�N�֐��`���ŕ�Ԃ��ꂽXList,YList
    
    '���͒l�̃`�F�b�N�y�яC��������������������������������������������������������
    '���͂��Z������(���[�N�V�[�g�֐�)�������ꍇ�̏���
    If IsObject(ArrayX1D) Then
        ArrayX1D = Application.Transpose(ArrayX1D.Value)
    End If
    If IsObject(ArrayY1D) Then
        ArrayY1D = Application.Transpose(ArrayY1D.Value)
    End If
    
    Dim StartNum%
    '�s��̊J�n�v�f��1�ɕύX�i�v�Z���₷������j
    StartNum = LBound(ArrayX1D, 1) '���͔z��̗v�f�̊J�n�ԍ�������Ă����i�o�͒l�ɍ��킹�邽�߁j
    If LBound(ArrayX1D, 1) <> 1 Then
        ArrayX1D = Application.Transpose(Application.Transpose(ArrayX1D))
    End If
    If LBound(ArrayY1D, 1) <> 1 Then
        ArrayY1D = Application.Transpose(Application.Transpose(ArrayY1D))
    End If
    
    '�z��̎����`�F�b�N
    Dim JigenCheck1%, JigenCheck2%
    On Error Resume Next
    JigenCheck1 = UBound(ArrayX1D, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    JigenCheck2 = UBound(ArrayY1D, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    On Error GoTo 0
    
    '�z��̎�����2�Ȃ玟��1�ɂ���B��)�z��(1 to N,1 to 1)���z��(1 to N)
    If JigenCheck1 > 0 Then
        ArrayX1D = Application.Transpose(ArrayX1D)
    End If
    If JigenCheck2 > 0 Then
        ArrayY1D = Application.Transpose(ArrayY1D)
    End If
    
    '�v�Z����������������������������������������������������������
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    N = UBound(ArrayX1D, 1)
    Dim ArrayT1D#(), ArrayParaT1D#()
    
    'X,Y�̕�Ԃ̊�ƂȂ�z����쐬
    ReDim ArrayT1D(1 To N)
    For I = 1 To N
        '0�`1�𓙊Ԋu
        ArrayT1D(I) = (I - 1) / (N - 1)
    Next I
    
    '�o�͕�Ԉʒu�̊�ʒu
    If JigenCheck1 > 0 Then '�o�͒l�̌`�����͒l�ɍ��킹�邽�߂̏���
        ReDim ArrayParaT1D(StartNum To StartNum - 1 + BunkatuN + 1, 1 To 1)
        For I = 1 To BunkatuN + 1
            '0�`1�𓙊Ԋu
            ArrayParaT1D(StartNum + I - 1, 1) = (I - 1) / (BunkatuN)
        Next I
    Else
        ReDim ArrayParaT1D(StartNum To StartNum - 1 + BunkatuN + 1)
        For I = 1 To BunkatuN + 1
            '0�`1�𓙊Ԋu
            ArrayParaT1D(StartNum + I - 1) = (I - 1) / (BunkatuN)
        Next I
    End If
    
    Dim OutputArrayX1D, OutputArrayY1D
    OutputArrayX1D = SplineByArrayX1DFast(ArrayT1D, ArrayX1D, ArrayParaT1D, PointCount)
    OutputArrayY1D = SplineByArrayX1DFast(ArrayT1D, ArrayY1D, ArrayParaT1D, PointCount)
    
    '�o��
    Dim Output(1 To 2)
    Output(1) = OutputArrayX1D
    Output(2) = OutputArrayY1D
    
    SplineParaFast = Output
    
End Function

Private Function SplineByArrayX1DFast(ByVal ArrayX1D, ByVal ArrayY1D, ByVal InputArrayX1D, PointCount&)
 '�X�v���C����Ԍv�Z���s��
 '�������Čv�Z���邱�Ƃō���������

'����
'HairetuX     �E�E�E��Ԃ̑ΏۂƂȂ�X�̒l���i�[���ꂽ�z��
'HairetuY     �E�E�E��Ԃ̑ΏۂƂȂ�Y�̒l���i�[���ꂽ�z��
'InputArrayX1D�E�E�E��ԈʒuX���i�[���ꂽ�z��
'PointCount   �E�E�E��������ۂ̓_��

'�Ԃ�l
'���͔z��InputArrayX1D�ɑ΂����Ԓl�̔z��
        
    '���͒l�̃`�F�b�N�y�яC��������������������������������������������������������
    '���͂��Z������(���[�N�V�[�g�֐�)�������ꍇ�̏���
    Dim RangeNaraTrue As Boolean
    RangeNaraTrue = False
    If IsObject(ArrayX1D) Then
        ArrayX1D = Application.Transpose(ArrayX1D.Value)
        RangeNaraTrue = True
    End If
    If IsObject(ArrayY1D) Then
        ArrayY1D = Application.Transpose(ArrayY1D.Value)
    End If
    If IsObject(InputArrayX1D) Then
        InputArrayX1D = Application.Transpose(InputArrayX1D.Value)
    End If
    
    Dim StartNum%
    '�s��̊J�n�v�f��1�ɕύX�i�v�Z���₷������j
    If LBound(ArrayX1D, 1) <> 1 Then
        ArrayX1D = Application.Transpose(Application.Transpose(ArrayX1D))
    End If
    If LBound(ArrayY1D, 1) <> 1 Then
        ArrayY1D = Application.Transpose(Application.Transpose(ArrayY1D))
    End If
    StartNum = LBound(InputArrayX1D, 1) 'InputArrayX1D�̊J�n�v�f�ԍ�������Ă����i�o�͒l�����킹�邽�߁j
    If LBound(InputArrayX1D, 1) <> 1 Then
        InputArrayX1D = Application.Transpose(Application.Transpose(InputArrayX1D))
    End If
    
    '�z��̎����`�F�b�N
    Dim JigenCheck1%, JigenCheck2%, JigenCheck3%
    On Error Resume Next
    JigenCheck1 = UBound(ArrayX1D, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    JigenCheck2 = UBound(ArrayY1D, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    JigenCheck3 = UBound(InputArrayX1D, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    On Error GoTo 0
    
    '�z��̎�����2�Ȃ玟��1�ɂ���B��)�z��(1 to N,1 to 1)���z��(1 to N)
    If JigenCheck1 > 0 Then
        ArrayX1D = Application.Transpose(ArrayX1D)
    End If
    If JigenCheck2 > 0 Then
        ArrayY1D = Application.Transpose(ArrayY1D)
    End If
    If JigenCheck3 > 0 Then
        InputArrayX1D = Application.Transpose(InputArrayX1D)
    End If

    '�v�Z����������������������������������������������������������
    Dim SplitArrayList
    SplitArrayList = �X�v���C����ԍ������p�ɕ�������(ArrayX1D, ArrayY1D, InputArrayX1D, PointCount)
        
    Dim TmpXList, TmpYList, TmpPointList
    Dim Output '�o�͒l�i�[�ϐ�
    Dim TmpSplineList
    Dim I&, J&, II&, JJ&, N&, M&, K&
    N = UBound(SplitArrayList, 1)
    K = 0
    For I = 1 To N
        TmpXList = SplitArrayList(I, 1)
        TmpYList = SplitArrayList(I, 2)
        TmpPointList = SplitArrayList(I, 3)
        If IsEmpty(TmpPointList) = False Then
            TmpSplineList = SplineByArrayX1D(TmpXList, TmpYList, TmpPointList)
            K = K + 1
            If K = 1 Then
                Output = TmpSplineList
            Else
                Output = UnionArray1D(Output, TmpSplineList)
            End If
        End If
    Next
    
    SplineByArrayX1DFast = Output
    
End Function

Private Function �X�v���C����ԍ������p�ɕ�������(ByVal ArrayX1D, ByVal ArrayY1D, ByVal CalPoint1D, PointCount&)
'�X�v���C����ԍ������p�ɕ�������
'20211009

'����
'ArrayX1D  �E�E�E��Ԍ���X���W���X�g
'ArrayY1D  �E�E�E��Ԍ���Y���W���X�g
'CalPoint1D�E�E�E��Ԉʒu��X���W���X�g
'PointCount�E�E�E������̈�̕����̓_��

    Dim I&, J&, II&, JJ&, N&, M&, K&
    N = UBound(ArrayX1D, 1)
    Dim PointN&
    PointN = UBound(CalPoint1D, 1)
    
    Dim Output '�o�͒l�i�[�ϐ�
    ReDim Output(1 To N, 1 To 3) '1:��Ԍ�X���W���X�g,2:��Ԍ�Y���W���X�g,3:��ԈʒuX���W���X�g
    'N�͂Ƃ肠�����̍ő�ŁA��Ŕz����k������
    
    Dim TmpXList, TmpYList, TmpPointList, TmpInterXList
    Dim StartNum&, EndNum& '���������Ԍ����W�̊J�n�ʒu�ƏI���ʒu
    Dim InterStartNum&, InterEndNum& '�������ꂽ��Ԍ����W�Ŏ��ۂ̕�Ԕ͈͂̊J�n�ʒu�ƏI���ʒu
    
    K = 0
    Do
        K = K + 1
        StartNum = (K - 1) * PointCount - 2
        EndNum = StartNum + PointCount + 2
        If StartNum <= 1 Then
            InterStartNum = 1
            StartNum = 1
        Else
            InterStartNum = StartNum + 1
        End If
        
        If EndNum >= N Then
            InterEndNum = N
            EndNum = N
        Else
            InterEndNum = EndNum - 1
        End If
        
        TmpXList = ExtractArray1D(ArrayX1D, StartNum, EndNum)
        TmpYList = ExtractArray1D(ArrayY1D, StartNum, EndNum)
        TmpInterXList = ExtractArray1D(ArrayX1D, InterStartNum, InterEndNum)
        TmpPointList = ExtractByRangeArray1D(CalPoint1D, TmpInterXList)
        
        Output(K, 1) = TmpXList
        Output(K, 2) = TmpYList
        Output(K, 3) = TmpPointList
        
        If EndNum = N Then
            Exit Do
        End If
    Loop
    
    '�o�͂���i�[�z��͈̔͒���
    Output = ExtractArray(Output, 1, 1, K, 3)
    
    '����������Ԉʒu�ŏd��������̂�����
    N = UBound(Output, 1)
    Dim TmpList1, TmpList2
    For I = 2 To N
        TmpList1 = Output(I - 1, 3)
        TmpList2 = Output(I, 3)
        If IsEmpty(TmpList1) = False And IsEmpty(TmpList2) = False Then
            If TmpList1(UBound(TmpList1, 1)) = TmpList2(1) Then '�Ō�̗v�f�ƍŏ��̗v�f���r����
                If UBound(TmpList2, 1) = 1 Then
                    TmpList2 = Empty
                Else
                    TmpList2 = ExtractArray1D(TmpList2, 2, UBound(TmpList2, 1))
                End If
                Output(I, 3) = TmpList2
            End If
        End If
    Next
    
    �X�v���C����ԍ������p�ɕ������� = Output
    
End Function

Private Function ExtractByRangeArray1D(InputArray1D, RangeArray1D)
'�ꎟ���z��̎w��͈͂𒊏o����B
'�w��͈͂�RangeArray1D�Ŏw�肷��B
'20211009

'����
'InputArray1D�E�E�E���o���̈ꎟ���z��
'RangeArray1D�E�E�E���o����͈͂��w�肷��ꎟ���z��

'��
'InputArray1D = (1,2,3,4,5,6,7,8,9,10)
'RangeArray1D = (3,4,7)
'�o�� = (3,4,5,6,7)

    '�����`�F�b�N
    Call CheckArray1D(InputArray1D, "InputArray1D")
    Call CheckArray1DStart1(InputArray1D, "InputArray1D")
    Call CheckArray1D(RangeArray1D, "RangeArray1D")
    Call CheckArray1DStart1(RangeArray1D, "RangeArray1D")
    
    Dim I&, J&, II&, JJ&, N&, M&, K&
    
    
    '�w��͈͂̍ŏ��A�ő���擾
    Dim MinNum#, MaxNum#
    MinNum = WorksheetFunction.Min(RangeArray1D)
    MaxNum = WorksheetFunction.Max(RangeArray1D)
    
    '���o�͈͂̊J�n�ʒu�A�I���ʒu���v�Z
    Dim StartNum&, EndNum&
    StartNum = 0
    EndNum = 0
    N = UBound(InputArray1D, 1)
    For I = 1 To N
        If InputArray1D(I) >= MinNum Then
            StartNum = I
            Exit For
        End If
    Next
    
    If StartNum = 0 Then
        '���o�͈͂Ȃ���Empty��Ԃ�
        Exit Function
    End If
    
    For I = StartNum To N
        If InputArray1D(I) > MaxNum Then
            EndNum = I - 1
            Exit For
        End If
    Next
    
    If EndNum = 0 Then
        '�I���ʒu��������Ȃ��ꍇ�͏I���܂őS���܂�
        EndNum = N
    End If
    
    '�͈͒��o
    Dim Output '�o�͒l�i�[�ϐ�
    Output = ExtractArray1D(InputArray1D, StartNum, EndNum)
    
    '�o��
    ExtractByRangeArray1D = Output
    
End Function

Private Sub CheckArray1D(InputArray, Optional HairetuName$ = "�z��")
'���͔z��1�����z�񂩂ǂ����`�F�b�N����
'20210804

    Dim Dummy%
    On Error Resume Next
    Dummy = UBound(InputArray, 2)
    On Error GoTo 0
    If Dummy <> 0 Then
        MsgBox (HairetuName & "��1�����z�����͂��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub

Private Sub CheckArray1DStart1(InputArray, Optional HairetuName$ = "�z��")
'����1�����z��̊J�n�ԍ���1���ǂ����`�F�b�N����
'20210804

    If LBound(InputArray, 1) <> 1 Then
        MsgBox (HairetuName & "�̊J�n�v�f�ԍ���1�ɂ��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub

Private Function ExtractArray1D(Array1D, StartNum&, EndNum&)
'�ꎟ���z��̎w��͈͂�z��Ƃ��Ē��o����
'20211009

'����
'Array1D �E�E�E�ꎟ���z��
'StartNum�E�E�E���o�͈͂̊J�n�ԍ�
'EndNum  �E�E�E���o�͈͂̏I���ԍ�
                                   
    '�����`�F�b�N
    Call CheckArray1D(Array1D, "Array1D")
    Call CheckArray1DStart1(Array1D, "Array1D")
    
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    N = UBound(Array1D, 1) '�v�f��
    
    If StartNum > EndNum Then
        MsgBox ("���o�͈͂̊J�n�ʒu�uStartNum�v�́A�I���ʒu�uEndNum�v�ȉ��łȂ���΂Ȃ�܂���")
        Stop
        Exit Function
    ElseIf StartNum < 1 Then
        MsgBox ("���o�͈͂̊J�n�ʒu�uStartNum�v��1�ȏ�̒l�����Ă�������")
        Stop
        Exit Function
    ElseIf EndNum > N Then
        MsgBox ("���o�͈͂̏I���s�uEndNum�v�͒��o���̈ꎟ���z��̗v�f��" & N & "�ȉ��̒l�����Ă�������")
        Stop
        Exit Function
    End If
    
    '����
    Dim Output
    ReDim Output(1 To EndNum - StartNum + 1)
    
    For I = StartNum To EndNum
        Output(I - StartNum + 1) = Array1D(I)
    Next I
    
    '�o��
    ExtractArray1D = Output
    
End Function

Private Function ExtractArray(Array2D, StartRow&, StartCol&, EndRow&, EndCol&)
'�񎟌��z��̎w��͈͂�z��Ƃ��Ē��o����
'20210917

'����
'Array2D �E�E�E�񎟌��z��
'StartRow�E�E�E���o�͈͂̊J�n�s�ԍ�
'StartCol�E�E�E���o�͈͂̊J�n��ԍ�
'EndRow  �E�E�E���o�͈͂̏I���s�ԍ�
'EndCol  �E�E�E���o�͈͂̏I����ԍ�
                                   
    '�����`�F�b�N
    Call CheckArray2D(Array2D, "Array2D")
    Call CheckArray2DStart1(Array2D, "Array2D")
    
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    N = UBound(Array2D, 1) '�s��
    M = UBound(Array2D, 2) '��
    
    If StartRow > EndRow Then
        MsgBox ("���o�͈͂̊J�n�s�uStartRow�v�́A�I���s�uEndRow�v�ȉ��łȂ���΂Ȃ�܂���")
        Stop
        End
    ElseIf StartCol > EndCol Then
        MsgBox ("���o�͈͂̊J�n��uStartCol�v�́A�I����uEndCol�v�ȉ��łȂ���΂Ȃ�܂���")
        Stop
        End
    ElseIf StartRow < 1 Then
        MsgBox ("���o�͈͂̊J�n�s�uStartRow�v��1�ȏ�̒l�����Ă�������")
        Stop
        End
    ElseIf StartCol < 1 Then
        MsgBox ("���o�͈͂̊J�n��uStartCol�v��1�ȏ�̒l�����Ă�������")
        Stop
        End
    ElseIf EndRow > N Then
        MsgBox ("���o�͈͂̏I���s�uStartRow�v�͒��o���̓񎟌��z��̍s��" & N & "�ȉ��̒l�����Ă�������")
        Stop
        End
    ElseIf EndCol > M Then
        MsgBox ("���o�͈͂̏I����uStartCol�v�͒��o���̓񎟌��z��̗�" & M & "�ȉ��̒l�����Ă�������")
        Stop
        End
    End If
    
    '����
    Dim Output
    ReDim Output(1 To EndRow - StartRow + 1, 1 To EndCol - StartCol + 1)
    
    For I = StartRow To EndRow
        For J = StartCol To EndCol
            Output(I - StartRow + 1, J - StartCol + 1) = Array2D(I, J)
        Next J
    Next I
    
    '�o��
    ExtractArray = Output
    
End Function

Private Sub CheckArray2D(InputArray, Optional HairetuName$ = "�z��")
'���͔z��2�����z�񂩂ǂ����`�F�b�N����
'20210804

    Dim Dummy2%, Dummy3%
    On Error Resume Next
    Dummy2 = UBound(InputArray, 2)
    Dummy3 = UBound(InputArray, 3)
    On Error GoTo 0
    If Dummy2 = 0 Or Dummy3 <> 0 Then
        MsgBox (HairetuName & "��2�����z�����͂��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub

Private Sub CheckArray2DStart1(InputArray, Optional HairetuName$ = "�z��")
'����2�����z��̊J�n�ԍ���1���ǂ����`�F�b�N����
'20210804

    If LBound(InputArray, 1) <> 1 Or LBound(InputArray, 2) <> 1 Then
        MsgBox (HairetuName & "�̊J�n�v�f�ԍ���1�ɂ��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub

Private Function SplineByArrayX1D(ByVal ArrayX1D, ByVal ArrayY1D, ByVal InputArrayX1D)
    '�X�v���C����Ԍv�Z���s��
    '���o�͒l�̐�����
    '���͔z��InputArrayX1D�ɑ΂����Ԓl�̔z��YList
    
    '�����͒l�̐�����
    'HairetuX�F��Ԃ̑ΏۂƂȂ�X�̒l���i�[���ꂽ�z��
    'HairetuY�F��Ԃ̑ΏۂƂȂ�Y�̒l���i�[���ꂽ�z��
    'InputArrayX1D:��ԈʒuX���i�[���ꂽ�z��

    '���͒l�̃`�F�b�N�y�яC��������������������������������������������������������
    '���͂��Z������(���[�N�V�[�g�֐�)�������ꍇ�̏���
    Dim RangeNaraTrue As Boolean
    RangeNaraTrue = False
    If IsObject(ArrayX1D) Then
        ArrayX1D = Application.Transpose(ArrayX1D.Value)
        RangeNaraTrue = True
    End If
    If IsObject(ArrayY1D) Then
        ArrayY1D = Application.Transpose(ArrayY1D.Value)
    End If
    If IsObject(InputArrayX1D) Then
        InputArrayX1D = Application.Transpose(InputArrayX1D.Value)
    End If
    
    Dim StartNum%
    '�s��̊J�n�v�f��1�ɕύX�i�v�Z���₷������j
    If LBound(ArrayX1D, 1) <> 1 Then
        ArrayX1D = Application.Transpose(Application.Transpose(ArrayX1D))
    End If
    If LBound(ArrayY1D, 1) <> 1 Then
        ArrayY1D = Application.Transpose(Application.Transpose(ArrayY1D))
    End If
    StartNum = LBound(InputArrayX1D, 1) 'InputArrayX1D�̊J�n�v�f�ԍ�������Ă����i�o�͒l�����킹�邽�߁j
    If LBound(InputArrayX1D, 1) <> 1 Then
        InputArrayX1D = Application.Transpose(Application.Transpose(InputArrayX1D))
    End If
    
    '�z��̎����`�F�b�N
    Dim JigenCheck1%, JigenCheck2%, JigenCheck3%
    On Error Resume Next
    JigenCheck1 = UBound(ArrayX1D, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    JigenCheck2 = UBound(ArrayY1D, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    JigenCheck3 = UBound(InputArrayX1D, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    On Error GoTo 0
    
    '�z��̎�����2�Ȃ玟��1�ɂ���B��)�z��(1 to N,1 to 1)���z��(1 to N)
    If JigenCheck1 > 0 Then
        ArrayX1D = Application.Transpose(ArrayX1D)
    End If
    If JigenCheck2 > 0 Then
        ArrayY1D = Application.Transpose(ArrayY1D)
    End If
    If JigenCheck3 > 0 Then
        InputArrayX1D = Application.Transpose(InputArrayX1D)
    End If

    '�v�Z����������������������������������������������������������
    Dim A, B, C, D
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    
    '�X�v���C���v�Z�p�̊e�W�����v�Z����B�Q�Ɠn����A,B,C,D�Ɋi�[
    Dim Dummy
    Dummy = SplineKeisu(ArrayX1D, ArrayY1D)
    A = Dummy(1)
    B = Dummy(2)
    C = Dummy(3)
    D = Dummy(4)
    
    Dim SotoNaraTrue As Boolean
    N = UBound(ArrayX1D, 1) '��ԑΏۂ̗v�f��
    
    Dim OutputArrayY1D#() '�o�͂���Y�̊i�[
    Dim NX%
    NX = UBound(InputArrayX1D, 1) '��Ԉʒu�̌�
    ReDim OutputArrayY1D(1 To NX)
    Dim TmpX#, TmpY#
    
    For J = 1 To NX
        TmpX = InputArrayX1D(J)
        SotoNaraTrue = False
        For I = 1 To N - 1
            If ArrayX1D(I) < ArrayX1D(I + 1) Then 'X���P�������̏ꍇ
                If I = 1 And ArrayX1D(1) > TmpX Then '�͈͂ɓ���Ȃ��Ƃ�(�J�n�_���O)
                    TmpY = ArrayY1D(1)
                    SotoNaraTrue = True
                    Exit For
                
                ElseIf I = N - 1 And ArrayX1D(I + 1) <= TmpX Then '�͈͂ɓ���Ȃ��Ƃ�(�I���_����)
                    TmpY = ArrayY1D(N)
                    SotoNaraTrue = True
                    Exit For
                    
                ElseIf ArrayX1D(I) <= TmpX And ArrayX1D(I + 1) > TmpX Then '�͈͓�
                    K = I: Exit For
                
                End If
            Else 'X���P�������̏ꍇ
            
                If I = 1 And ArrayX1D(1) < TmpX Then '�͈͂ɓ���Ȃ��Ƃ�(�J�n�_���O)
                    TmpY = ArrayY1D(1)
                    SotoNaraTrue = True
                    Exit For
                
                ElseIf I = N - 1 And ArrayX1D(I + 1) >= TmpX Then '�͈͂ɓ���Ȃ��Ƃ�(�I���_����)
                    TmpY = ArrayY1D(N)
                    SotoNaraTrue = True
                    Exit For
                    
                ElseIf ArrayX1D(I + 1) < TmpX And ArrayX1D(I) >= TmpX Then '�͈͓�
                    K = I: Exit For
                
                End If
            
            End If
        Next I
        
        If SotoNaraTrue = False Then
            TmpY = A(K) + B(K) * (TmpX - ArrayX1D(K)) + C(K) * (TmpX - ArrayX1D(K)) ^ 2 + D(K) * (TmpX - ArrayX1D(K)) ^ 3
        End If
        
        OutputArrayY1D(J) = TmpY
        
    Next J
    
    '�o�́�����������������������������������������������������
    Dim Output
    
    '�o�͂���z�����͂����z��InputArrayX1D�̌`��ɍ��킹��
    If JigenCheck3 = 1 Then '���͂�InputArrayX1D���񎟌��z��
        ReDim Output(StartNum To StartNum + NX - 1, 1 To 1)
        For I = 1 To NX
            Output(StartNum + I - 1, 1) = OutputArrayY1D(I)
        Next I
    Else
        If StartNum = 1 Then
            Output = OutputArrayY1D
        Else
            ReDim Output(StartNum To StartNum + NX - 1)
            For I = 1 To NX
                Output(StartNum + I - 1) = OutputArrayY1D(I)
            Next I
        End If
    End If
    
    If RangeNaraTrue Then
        '���[�N�V�[�g�֐��̏ꍇ
        SplineByArrayX1D = Application.Transpose(Output)
    Else
        'VBA��ł̏����̏ꍇ
        SplineByArrayX1D = Output
    End If
    
End Function

Private Function SplineKeisu(ByVal ArrayX1D, ByVal ArrayY1D)

    '�Q�l�Fhttp://www5d.biglobe.ne.jp/stssk/maze/spline.html
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    Dim A, B, C, D
    N = UBound(ArrayX1D, 1)
    ReDim A(1 To N)
    ReDim B(1 To N)
    ReDim D(1 To N)
    
    Dim h#()
    Dim ArrayL2D#() '���ӂ̔z�� �v�f��(1 to N,1 to N)
    Dim ArrayR1D#() '�E�ӂ̔z�� �v�f��(1 to N,1 to 1)
    Dim ArrayLm2D#() '���ӂ̔z��̋t�s�� �v�f��(1 to N,1 to N)
    
    ReDim h(1 To N - 1)
    ReDim ArrayL2D(1 To N, 1 To N)
    ReDim ArrayR1D(1 To N, 1 To 1)
    
    'hi = xi+1 - x
    For I = 1 To N - 1
        h(I) = ArrayX1D(I + 1) - ArrayX1D(I)
    Next I
    
    'di = yi
    For I = 1 To N
        A(I) = ArrayY1D(I)
    Next I
    
    '�E�ӂ̔z��̌v�Z
    For I = 1 To N
        If I = 1 Or I = N Then
            ArrayR1D(I, 1) = 0
        Else
            ArrayR1D(I, 1) = 3 * (ArrayY1D(I + 1) - ArrayY1D(I)) / h(I) - 3 * (ArrayY1D(I) - ArrayY1D(I - 1)) / h(I - 1)
        End If
    Next I
    
    '���ӂ̔z��̌v�Z
    For I = 1 To N
        If I = 1 Then
            ArrayL2D(I, 1) = 1
        ElseIf I = N Then
            ArrayL2D(N, N) = 1
        Else
            ArrayL2D(I - 1, I) = h(I - 1)
            ArrayL2D(I, I) = 2 * (h(I) + h(I - 1))
            ArrayL2D(I + 1, I) = h(I)
        End If
    Next I
    
    '���ӂ̔z��̋t�s��
    ArrayLm2D = F_Minverse(ArrayL2D)
    
    'C�̔z������߂�
    C = F_MMult(ArrayLm2D, ArrayR1D)
    C = Application.Transpose(C)
    
    'B�̔z������߂�
    For I = 1 To N - 1
        B(I) = (A(I + 1) - A(I)) / h(I) - h(I) * (C(I + 1) + 2 * C(I)) / 3
    Next I
    
    'D�̔z������߂�
    For I = 1 To N - 1
        D(I) = (C(I + 1) - C(I)) / (3 * h(I))
    Next I
    
    '�o��
    Dim Output(1 To 4)
    Output(1) = A
    Output(2) = B
    Output(3) = C
    Output(4) = D
    
    SplineKeisu = Output

End Function

Private Function F_MMult(ByVal Matrix1, ByVal Matrix2)
    'F_MMult(Matrix1, Matrix2)
    'F_MMult(�z��@,�z��A)
    '�s��̐ς��v�Z
    '20180213����
    '20210603����
    
    '���͒l�̃`�F�b�N�ƏC��������������������������������������������������������
    '�z��̎����`�F�b�N
    Dim JigenCheck1%, JigenCheck2%
    On Error Resume Next
    JigenCheck1 = UBound(Matrix1, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    JigenCheck2 = UBound(Matrix2, 2) '�z��̎�����1�Ȃ�G���[�ƂȂ�
    On Error GoTo 0
    
    '�z��̎�����1�Ȃ玟��2�ɂ���B��)�z��(1 to N)���z��(1 to N,1 to 1)
    If IsEmpty(JigenCheck1) Then
        Matrix1 = Application.Transpose(Matrix1)
    End If
    If IsEmpty(JigenCheck2) Then
        Matrix2 = Application.Transpose(Matrix2)
    End If
    
    '�s��̊J�n�v�f��1�ɕύX�i�v�Z���₷������j
    If UBound(Matrix1, 1) = 0 Or UBound(Matrix1, 2) = 0 Then
        Matrix1 = Application.Transpose(Application.Transpose(Matrix1))
    End If
    If UBound(Matrix2, 1) = 0 Or UBound(Matrix2, 2) = 0 Then
        Matrix2 = Application.Transpose(Application.Transpose(Matrix2))
    End If
    
    '���͒l�̃`�F�b�N
    If UBound(Matrix1, 2) <> UBound(Matrix2, 1) Then
        MsgBox ("�z��1�̗񐔂Ɣz��2�̍s������v���܂���B" & vbLf & _
               "(�o��) = (�z��1)(�z��2)")
        Stop
        End
    End If
    
    '�v�Z����������������������������������������������������������
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    Dim M2%
    Dim Output#() '�o�͂���z��
    N = UBound(Matrix1, 1) '�z��1�̍s��
    M = UBound(Matrix1, 2) '�z��1�̗�
    M2 = UBound(Matrix2, 2) '�z��2�̗�
    
    ReDim Output(1 To N, 1 To M2)
    
    For I = 1 To N '�e�s
        For J = 1 To M2 '�e��
            For K = 1 To M '(�z��1��I�s)��(�z��2��J��)���|�����킹��
                Output(I, J) = Output(I, J) + Matrix1(I, K) * Matrix2(K, J)
            Next K
        Next J
    Next I
    
    '�o�́�����������������������������������������������������
    F_MMult = Output
    
End Function

Private Function F_Minverse(ByVal Matrix)
    '20210603����
    'F_Minverse(input_M)
    'F_Minverse(�z��)
    '�]���q�s���p���ċt�s����v�Z
    
    '���͒l�`�F�b�N�y�яC��������������������������������������������������������
    '�s��̊J�n�v�f��1�ɕύX�i�v�Z���₷������j
    If LBound(Matrix, 1) <> 1 Or LBound(Matrix, 2) <> 1 Then
        Matrix = Application.Transpose(Application.Transpose(Matrix))
    End If
    
    '���͒l�̃`�F�b�N
    Call �����s�񂩃`�F�b�N(Matrix)
    
    '�v�Z����������������������������������������������������������
    Dim I%, J%, K%, M%, M2%, N% '�����グ�p(Integer�^)
    N = UBound(Matrix, 1)
    Dim Output#()
    ReDim Output(1 To N, 1 To N)
    
    Dim detM# '�s�񎮂̒l���i�[
    detM = F_MDeterm(Matrix) '�s�񎮂����߂�
    
    Dim Mjyokyo '�w��̗�E�s�����������z����i�[
    
    For I = 1 To N '�e��
        For J = 1 To N '�e�s
            
            'I��,J�s����������
            Mjyokyo = F_Mjyokyo(Matrix, J, I)
            
            'I��,J�s�̗]���q�����߂ďo�͂���t�s��Ɋi�[
            Output(I, J) = F_MDeterm(Mjyokyo) * (-1) ^ (I + J) / detM
    
        Next J
    Next I
    
    '�o�́�����������������������������������������������������
    F_Minverse = Output
    
End Function

Private Sub �����s�񂩃`�F�b�N(Matrix)
    '20210603�ǉ�
    
    If UBound(Matrix, 1) <> UBound(Matrix, 2) Then
        MsgBox ("�����s�����͂��Ă�������" & vbLf & _
                "���͂��ꂽ�z��̗v�f����" & "�u" & _
                UBound(Matrix, 1) & "�~" & UBound(Matrix, 2) & "�v" & "�ł�")
        Stop
        End
    End If

End Sub

Private Function F_MDeterm(Matrix)
    '20210603����
    'F_MDeterm(Matrix)
    'F_MDeterm(�z��)
    '�s�񎮂��v�Z
    
    '���͒l�`�F�b�N�y�яC��������������������������������������������������������
    '�s��̊J�n�v�f��1�ɕύX�i�v�Z���₷������j
    If LBound(Matrix, 1) <> 1 Or LBound(Matrix, 2) <> 1 Then
        Matrix = Application.Transpose(Application.Transpose(Matrix))
    End If
    
    '���͒l�̃`�F�b�N
    Call �����s�񂩃`�F�b�N(Matrix)
    
    '�v�Z����������������������������������������������������������
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    N = UBound(Matrix, 1)
    
    Dim Matrix2 '�|���o�����s���s��
    Matrix2 = Matrix
    
    For I = 1 To N '�e��
        For J = I To N '�|���o�����̍s�̒T��
            If Matrix2(J, I) <> 0 Then
                K = J '�|���o�����̍s
                Exit For
            End If
            
            If J = N And Matrix2(J, I) = 0 Then '�|���o�����̒l���S��0�Ȃ�s�񎮂̒l��0
                F_MDeterm = 0
                Exit Function
            End If
            
        Next J
        
        If K <> I Then '(I��,I�s)�ȊO�ő|���o���ƂȂ�ꍇ�͍s�����ւ�
            Matrix2 = F_Mgyoirekae(Matrix2, I, K)
        End If
        
        '�|���o��
        Matrix2 = F_Mgyohakidasi(Matrix2, I, I)
              
    Next I
    
    
    '�s�񎮂̌v�Z
    Dim Output#
    Output = 1
    
    For I = 1 To N '�e(I��,I�s)���|�����킹�Ă���
        Output = Output * Matrix2(I, I)
    Next I
    
    '�o�́�����������������������������������������������������
    F_MDeterm = Output
    
End Function

Private Function F_Mgyoirekae(Matrix, Row1%, Row2%)
    '20210603����
    'F_Mgyoirekae(Matrix, Row1, Row2)
    'F_Mgyoirekae(�z��,�w��s�ԍ��@,�w��s�ԍ��A)
    '�s��Matrix�̇@�s�ƇA�s�����ւ���
    
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    Dim Output
    
    Output = Matrix
    M = UBound(Matrix, 2) '�񐔎擾
    
    For I = 1 To M
        Output(Row2, I) = Matrix(Row1, I)
        Output(Row1, I) = Matrix(Row2, I)
    Next I
    
    F_Mgyoirekae = Output
End Function

Private Function F_Mgyohakidasi(Matrix, Row%, Col%)
    '20210603����
    'F_Mgyohakidasi(Matrix, Row, Col)
    'F_Mgyohakidasi(�z��,�w��s,�w���)
    '�s��Matrix��Row�s�Col��̒l�Ŋe�s��|���o��
    
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    Dim Output
    
    Output = Matrix
    N = UBound(Output, 1) '�s���擾
    
    Dim Hakidasi '�|���o�����̍s
    Dim X# '�|���o�����̒l
    Dim Y#
    ReDim Hakidasi(1 To N)
    X = Matrix(Row, Col)
    
    For I = 1 To N '�|���o������1�s���쐬
        Hakidasi(I) = Matrix(Row, I)
    Next I
    
    
    For I = 1 To N '�e�s
        If I = Row Then
            '�|���o�����̍s�̏ꍇ�͂��̂܂�
            For J = 1 To N
                Output(I, J) = Matrix(I, J)
            Next J
        
        Else
            '�|���o�����̍s�ȊO�̏ꍇ�͑|���o��
            Y = Matrix(I, Col) '�|���o����̗�̒l
            For J = 1 To N
                Output(I, J) = Matrix(I, J) - Hakidasi(J) * Y / X
            Next J
        End If
    
    Next I
    
    F_Mgyohakidasi = Output
    
End Function

Private Function F_Mjyokyo(Matrix, Row%, Col%)
    '20210603����
    'F_Mjyokyo(Matrix, Row, Col)
    'F_Mjyokyo(�z��,�w��s,�w���)
    '�s��Matrix��Row�s�ACol������������s���Ԃ�
    
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    Dim Output '�w�肵���s�E���������̔z��
    
    N = UBound(Matrix, 1) '�s���擾
    M = UBound(Matrix, 2) '�񐔎擾
    ReDim Output(1 To N - 1, 1 To M - 1)
    
    Dim I2%, J2%
    
    I2 = 0 '�s���������グ������
    For I = 1 To N
        If I = Row Then
            '�Ȃɂ����Ȃ�
        Else
            I2 = I2 + 1 '�s���������グ
            
            J2 = 0 '����������グ������
            For J = 1 To M
                If J = Col Then
                    '�Ȃɂ����Ȃ�
                Else
                    J2 = J2 + 1 '����������グ
                    Output(I2, J2) = Matrix(I, J)
                End If
            Next J
            
        End If
    Next I
    
    F_Mjyokyo = Output

End Function

Private Function UnionArray1D(UpperArray1D, LowerArray1D)
'�ꎟ���z�񓯎m����������1�̔z��Ƃ���B
'20210923

'UpperArray1D�E�E�E��Ɍ�������ꎟ���z��
'LowerArray1D�E�E�E���Ɍ�������ꎟ���z��

    '�����`�F�b�N
    Call CheckArray1D(UpperArray1D, "UpperArray1D")
    Call CheckArray1DStart1(UpperArray1D, "UpperArray1D")
    Call CheckArray1D(LowerArray1D, "LowerArray1D")
    Call CheckArray1DStart1(LowerArray1D, "LowerArray1D")
    
    '����
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    Dim N1&, N2&
    N1 = UBound(UpperArray1D, 1)
    N2 = UBound(LowerArray1D, 1)
    Dim Output
    ReDim Output(1 To N1 + N2)
    For I = 1 To N1
        Output(I) = UpperArray1D(I)
    Next I
    For I = 1 To N2
        Output(N1 + I) = LowerArray1D(I)
    Next I
    
    '�o��
    UnionArray1D = Output
    
End Function

Private Function DrawPolyLine(XYList, TargetSheet As Worksheet) As Shape
'XY���W����|�����C����`��
'�V�F�C�v���I�u�W�F�N�g�ϐ��Ƃ��ĕԂ�
'20210921

'����
'XYList         �E�E�EXY���W���������񎟌��z�� X�������E���� Y������������
'TargetSheet    �E�E�E��}�Ώۂ̃V�[�g

    Dim I%, Count%
    Count = UBound(XYList, 1)
    
    With TargetSheet.Shapes.BuildFreeform(msoEditingCorner, XYList(1, 1), XYList(1, 2))
        
        For I = 2 To Count
            .AddNodes msoSegmentLine, msoEditingAuto, XYList(I, 1), XYList(I, 2)
        Next I
        Set DrawPolyLine = .ConvertToShape
    End With
    
End Function

Private Function GetXYDocumentFromCursor(Optional ImmidiateShow As Boolean = True)
'���݃J�[�\���ʒu�̃h�L�������g���W�擾
'�J�[�\���ʒu�̃X�N���[�����W���A
'�J�[�\��������Ă���Z���̎l���̃X�N���[�����W�̊֌W������A
'�J�[�\��������Ă���Z���̎l���̃h�L�������g���W�����Ƃɕ�Ԃ��āA
'�J�[�\���ʒu�̃h�L�������g���W�����߂�B
'20211005

'����
'[ImmidiateShow]�E�E�E�C�~�f�B�G�C�g�E�B���h�E�Ɍv�Z���ʂȂǂ�\�����邩(�f�t�H���g��True)

'�Ԃ�l
'Output(1 to 2)�E�E�E1:�J�[�\���ʒu�̃h�L�������g���WX,2:�J�[�\���ʒu�̃h�L�������g���WY(Double�^)
'�J�[�\�����V�[�g���ɂȂ��ꍇ��Empty��Ԃ��B

'�Q�l�Fhttps://gist.github.com/furyutei/f0668f33d62ccac95d1643f15f19d99a?s=09#to-footnote-1

    Dim Win As Window
    Set Win = ActiveWindow
    
    '�J�[�\���̃X�N���[�����W�擾
    Dim Cursor As PointAPI, CursorScreenX#, CursorScreenY#
    Call GetCursorPos(Cursor)
    CursorScreenX = Cursor.X
    CursorScreenY = Cursor.Y
    
    '�J�[�\��������Ă���Z�����擾
    Dim CursorCell As Range, Dummy
    Set Dummy = Win.RangeFromPoint(CursorScreenX, CursorScreenY)
    If TypeName(Dummy) = "Range" Then
        Set CursorCell = Dummy
    Else
        '�J�[�\�����Z���ɏ���ĂȂ��̂ŏI��
        Exit Function
    End If
    
    '�l���̃X�N���[�����W���擾
    Dim X1Screen#, X2Screen#, Y1Screen#, Y2Screen# '�l���̃X�N���[�����W
    Dummy = GetXYCellScreenUpperLeft(CursorCell)
    If IsEmpty(Dummy) Then Exit Function
    X1Screen = Dummy(1)
    Y1Screen = Dummy(2)
    
    Dummy = GetXYCellScreenLowerRight(CursorCell)
    If IsEmpty(Dummy) Then Exit Function
    X2Screen = Dummy(1)
    Y2Screen = Dummy(2)
    
    '�l���̃h�L�������g���W�擾
    Dim X1Document#, X2Document#, Y1Document#, Y2Document# '�l���̃h�L�������g���W
    X1Document = CursorCell.Left
    X2Document = CursorCell.Left + CursorCell.Width
    Y1Document = CursorCell.Top
    Y2Document = CursorCell.Top + CursorCell.Height
    
    '�}�E�X�J�[�\���̃h�L�������g���W���ԂŌv�Z
    Dim CursorDocumentX#, CursorDocumentY#
    CursorDocumentX = X1Document + (X2Document - X1Document) * (CursorScreenX - X1Screen) / (X2Screen - X1Screen)
    CursorDocumentY = Y1Document + (Y2Document - Y1Document) * (CursorScreenY - Y1Screen) / (Y2Screen - Y1Screen)
        
    '�o��
    Dim Output#(1 To 2)
    Output(1) = CursorDocumentX
    Output(2) = CursorDocumentY
    
    GetXYDocumentFromCursor = Output
    
    '�m�F�\��
    If ImmidiateShow Then
        Debug.Print "�J�[�\���̏�����Z��", CursorCell.Address(False, False)
        Debug.Print "�J�[�\���X�N���[�����W", "CursorScreenX:" & CursorScreenX, "CursorScreenY:" & CursorScreenY
        Debug.Print "�J�[�\���h�L�������g���W", "CursorDocumentX:" & WorksheetFunction.Round(CursorDocumentX, 1), "CursorDocumentY:" & WorksheetFunction.Round(CursorDocumentY, 1)
        Debug.Print "�Z������X�N���[�����W", "X1Screen:" & X1Screen, , "Y1Screen:" & Y1Screen
        Debug.Print "�Z������h�L�������g���W", "X1Document:" & X1Document, "Y1Document:" & Y1Document
        Debug.Print "�Z���E���X�N���[�����W", "X2Screen:" & X2Screen, , "Y2Screen:" & Y2Screen
        Debug.Print "�Z���E���h�L�������g���W", "X2Document:" & X2Document, "Y2Document:" & Y2Document
    End If

End Function

Private Function GetXYCellScreenUpperLeft(TargetCell As Range)
'�w��Z���̍���̃X�N���[�����WXY���擾����B
'20211005

'����
'TargetCell�E�E�E�Ώۂ̃Z��(Range�^)

'�Ԃ�l
'Output(1 to 2)�E�E�E1:�Z������̃X�N���[�����WX,2;�Z������̃X�N���[�����WY(Double�^)

    '�Z�����\������Ă���Pane(�E�B���h�E�g�̌Œ���l�������\���G���A)
    Dim Pane As Pane
    Set Pane = GetPaneOfCell(TargetCell)
    If Pane Is Nothing Then Exit Function
       
    '�yPointsToScreenPixels�̒��ӎ����z
    '�y���z�ΏۃZ�����V�[�g��ŕ\������Ă��Ȃ��Ǝ擾�s�B�ꕔ�ł��\������Ă�����擾�\�B
    Dim Output#(1 To 2)
    Output(1) = Pane.PointsToScreenPixelsX(TargetCell.Left)
    Output(2) = Pane.PointsToScreenPixelsY(TargetCell.Top)
    
    GetXYCellScreenUpperLeft = Output
    
End Function

Private Function GetPaneOfCell(TargetCell As Range) As Pane
'�w��Z����Pane���擾����
'�E�B���h�E�g�Œ�A�E�B���h�E�����̐ݒ�ł��擾�ł���B
'�Q�l�Fhttp://www.asahi-net.or.jp/~ef2o-inue/vba_o/sub05_100_120.html
'20211006

'����
'TargetCell�E�E�E�Ώۂ̃Z��/Range�^

'�Ԃ�l
'�w��Z�����܂܂��Pane/Pane�^
'�w��Z�����\���͈͊O�Ȃ�Nothing
    
    Dim Win As Window
    Set Win = ActiveWindow
    
    Dim Output As Pane
    Dim I& '�����グ�p(Long�^)
    
    ' �E�B���h�E����������
    If Not Win.FreezePanes And Not Win.Split Then
        '�E�B���h�E�g�Œ�ł��E�B���h�E�����ł��Ȃ��ꍇ
        ' �\���ȊO�ɃZ��������ꍇ�͖���
        If Intersect(Win.VisibleRange, TargetCell) Is Nothing Then Exit Function
        Set Output = Win.Panes(1)
    Else ' ��������
        If Win.FreezePanes Then
            ' �E�B���h�E�g�Œ�̏ꍇ
            ' �ǂ̃E�B���h�E�ɑ����邩����
            For I = 1 To Win.Panes.Count
                If Not Intersect(Win.Panes(I).VisibleRange, TargetCell) Is Nothing Then
                    'Pane�̕\���͈͂Ɋ܂܂��ꍇ�͂���Pane���擾
                    Set Output = Win.Panes(I)
                    Exit For
                End If
            Next I
            
            '������Ȃ������ꍇ
            If Output Is Nothing Then Exit Function
        Else
            '�E�B���h�E�����̏ꍇ
            ' �E�B���h�E�����̓A�N�e�B�u�y�C���̂ݔ���
            If Not Intersect(Win.ActivePane.VisibleRange, TargetCell) Is Nothing Then
                Set Output = Win.ActivePane
            Else
                Exit Function
            End If
        End If
    End If
    
    '�o��
    Set GetPaneOfCell = Output
    
End Function

Private Function GetXYCellScreenLowerRight(TargetCell As Range)
'�w��Z���̉E���̃X�N���[�����WXY���擾����B
'20211005

'����
'TargetCell�E�E�E�Ώۂ̃Z��(Range�^)

'�Ԃ�l
'Output(1 to 2)�E�E�E1:�Z���E���̃X�N���[�����WX,2;�Z���E���̃X�N���[�����WY(Double�^)

    '�Z�����\������Ă���Pane(�E�B���h�E�g�̌Œ���l�������\���G���A)
    Dim Pane As Pane
    Set Pane = GetPaneOfCell(TargetCell)
    If Pane Is Nothing Then Exit Function
    
    '�yPointsToScreenPixels�̒��ӎ����z
    '�y���z�ΏۃZ�����V�[�g��ŕ\������Ă��Ȃ��Ǝ擾�s�B�ꕔ�ł��\������Ă�����擾�\�B
    Dim Output#(1 To 2)
    Output(1) = Pane.PointsToScreenPixelsX(TargetCell.Left + TargetCell.Width)
    Output(2) = Pane.PointsToScreenPixelsY(TargetCell.Top + TargetCell.Height)
    
    GetXYCellScreenLowerRight = Output
    
End Function
