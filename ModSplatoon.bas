Attribute VB_Name = "ModSplatoon"
Option Explicit

'Bloodstain                      ・・・元場所：VBAProject.Module1
'血痕の計算                      ・・・元場所：VBAProject.Module1
'SplineXYParaFast                ・・・元場所：FukamiAddins3.ModApproximate
'SplineParaFast                  ・・・元場所：FukamiAddins3.ModApproximate
'SplineByArrayX1DFast            ・・・元場所：FukamiAddins3.ModApproximate
'スプライン補間高速化用に分割処理・・・元場所：FukamiAddins3.ModApproximate
'ExtractByRangeArray1D           ・・・元場所：FukamiAddins3.ModApproximate
'CheckArray1D                    ・・・元場所：FukamiAddins3.ModArray
'CheckArray1DStart1              ・・・元場所：FukamiAddins3.ModArray
'ExtractArray1D                  ・・・元場所：FukamiAddins3.ModArray
'ExtractArray                    ・・・元場所：FukamiAddins3.ModArray
'CheckArray2D                    ・・・元場所：FukamiAddins3.ModArray
'CheckArray2DStart1              ・・・元場所：FukamiAddins3.ModArray
'SplineByArrayX1D                ・・・元場所：FukamiAddins3.ModApproximate
'SplineKeisu                     ・・・元場所：FukamiAddins3.ModApproximate
'F_MMult                         ・・・元場所：FukamiAddins3.ModMatrix
'F_Minverse                      ・・・元場所：FukamiAddins3.ModMatrix
'正方行列かチェック              ・・・元場所：FukamiAddins3.ModMatrix
'F_MDeterm                       ・・・元場所：FukamiAddins3.ModMatrix
'F_Mgyoirekae                    ・・・元場所：FukamiAddins3.ModMatrix
'F_Mgyohakidasi                  ・・・元場所：FukamiAddins3.ModMatrix
'F_Mjyokyo                       ・・・元場所：FukamiAddins3.ModMatrix
'UnionArray1D                    ・・・元場所：FukamiAddins3.ModArray
'DrawPolyLine                    ・・・元場所：FukamiAddins3.ModDrawShape
'GetXYDocumentFromCursor         ・・・元場所：FukamiAddins3.ModCursor
'GetXYCellScreenUpperLeft        ・・・元場所：FukamiAddins3.ModCursor
'GetPaneOfCell                   ・・・元場所：FukamiAddins3.ModCursor
'GetXYCellScreenLowerRight       ・・・元場所：FukamiAddins3.ModCursor

'------------------------------
Const Pi# = 3.141529
'------------------------------
'シート関数用近似、補間関数
'------------------------------
'配列の処理関係のプロシージャ
'------------------------------
'行列を使った計算
'代替関数
'------------------------------
'シェイプ作図関連モジュール
'20210914作成
'------------------------------
'※※※※※※※※※※※※※※※※※※※※※※※※※※※
'カーソルのスクリーン座標取得用
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
'カーソル位置に血が飛ぶ
'20211009

'引数
'TargetSheet・・・地を飛ばす対象のシート/Worksheet型

    'カーソル位置のドキュメント座標取得
    Dim CenterX#, CenterY#
    Dim Dummy
    On Error Resume Next 'スクリーン座標取得に失敗した場合
    Dummy = GetXYDocumentFromCursor
    CenterX = Dummy(1)
    CenterY = Dummy(2)
    On Error GoTo 0
    If CenterX = 0 Then
        Exit Sub
    End If
    
    Dim N&, I&
    Dim k_fai1#, k_fai2#, k_fai3#, r0#, kr#, p#
    
    N = 10 + Rnd() * 6 '血痕のツノの個数
    k_fai1 = 0.4 '血痕のツノの根っこの座標の角度の係数
    k_fai2 = 0.11 '血痕のツノのくびれの座標の角度の係数
    k_fai3 = 0.2 '血痕のツノの膨らみの座標の角度の係数
    r0 = 4 / 20 * N * 1.2   '血痕の核半径
    kr = 0.9 + 0.2 * Rnd() 'ツノの長さ係数。大きいほどツノが長くなる
    p = 0.3 '調整係数(隣通しのツノとの間隔のランダム調整)(大きいほど間隔が大きく変わる)
    
    '色リスト
    Dim ColorList&(1 To 6)
    ColorList(1) = RGB(0, 0, 255) '青
    ColorList(2) = RGB(231, 34, 231) '紫
    ColorList(3) = RGB(255, 124, 0) 'オレンジ
    ColorList(4) = RGB(0, 255, 255) '水色
    ColorList(5) = RGB(158, 255, 69) '黄緑
    ColorList(6) = RGB(255, 0, 148) '紫2
    
    Dim ColorNum&, InputColor&
    ColorNum = WorksheetFunction.RandBetween(1, 6)
    InputColor = ColorList(ColorNum)
    
    Call 血痕の計算(N, k_fai1, k_fai2, k_fai3, r0, kr, p, CenterX, CenterY, TargetSheet, InputColor)

End Sub

Private Sub 血痕の計算(N&, k_fai1#, k_fai2#, k_fai3#, r0#, kr#, p#, CenterX#, CenterY#, TargetSheet As Worksheet, Optional InputColor& = rgbRed)
'血痕の形状を計算して、指定位置に描画
'20211009

'N          ・・・血痕のツノの個数
'k_fai1     ・・・血痕のツノの根っこの座標の角度の係数
'k_fai2     ・・・血痕のツノのくびれの座標の角度の係数
'k_fai3     ・・・血痕のツノの膨らみの座標の角度の係数
'r0         ・・・血痕の核半径
'kr         ・・・ツノの長さ係数。大きいほどツノが長くなる
'p          ・・・調整係数(隣通しのツノとの間隔のランダム調整)(大きいほど間隔が大きく変わる)
'CenterX    ・・・血痕の中心X
'CenterY    ・・・血痕の中心Y
'TargetSheet・・・血を飛ばす対象のシート
'InputColor ・・・塗りつぶし色。デフォルトは赤

    Dim I&
    
    Dim ThetaList#(), ThetaDashList#()
    Dim dTheta#
    ReDim ThetaList(1 To N)
    ReDim ThetaDashList(1 To N)
    
    For I = 1 To N
        ThetaList(I) = 2 * Pi / N * I - Pi / N   'θi
        dTheta = p * Pi / N * (2 * Rnd() - 1)    'dθ
        ThetaDashList(I) = ThetaList(I) + dTheta 'θ'i
    Next I
    
    Dim Fai1#, Fai2#, Fai3#, FaiList#(), rList#()
    Dim dr#
    Fai1 = k_fai1 * Pi / N 'φ1
    Fai2 = k_fai2 * Pi / N 'φ1
    Fai3 = k_fai3 * Pi / N 'φ1
    
    ReDim FaiList(1 To N, 1 To 3) 'φi_1,φi_2,φi_3
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
    
    'スプライン補間で点を増やす
    Dim BunkatuN&
    BunkatuN = 1000
    XYList = SplineXYParaFast(XYList, BunkatuN, 4)
    
    '作図中心位置へ移動
    For I = 1 To UBound(XYList, 1)
        XYList(I, 1) = CenterX + XYList(I, 1) * 10
        XYList(I, 2) = CenterY + XYList(I, 2) * 10
    Next
    
    '出力
    Dim TmpShape As Shape
'    Application.ScreenUpdating = False
    Set TmpShape = DrawPolyLine(XYList, TargetSheet)
'    Set TmpShape = 曲線を作図する(XYList, CenterX, CenterY, 10)
'    Application.ScreenUpdating = True
    With TmpShape
        .Fill.ForeColor.RGB = InputColor
        .Line.ForeColor.RGB = InputColor
    End With
        
End Sub

Private Function SplineXYParaFast(ByVal ArrayXY2D, BunkatuN&, PointCount&)
'パラメトリック関数形式でスプライン補間を行う
'分割して計算を高速化する
'ArrayX,ArrayYがどちらも単調増加、単調減少でない場合に用いる。
    
'引数
'ArrayXY2D ・・・補間の対象となるX,Yの値が格納された配列
'ArrayXY2Dの1列目がX,2列目がYとなるようにする。
'BunkatuN  ・・・パラメトリック関数の分割個数（出力されるXList,YListの要素数は(分割個数+1)）
'PointCount・・・分割する際の点数
    
'返り値
'パラメトリック関数形式で補間されたXList,YListが格納されたXYList
'1列目がXList,2列目がYList
    
    '入力値のチェック及び修正'※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '入力がセルから(ワークシート関数)だった場合の処理
    If IsObject(ArrayXY2D) Then
        ArrayXY2D = ArrayXY2D.Value
    End If
        
    '行列の開始要素を1に変更（計算しやすいから）
    Dim StartNum%
    StartNum = LBound(ArrayXY2D) '入力配列の要素の開始番号を取っておく（出力値に合わせるため）
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
    
    '計算処理※※※※※※※※※※※※※※※※※※※※※※※※※※※
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
    
    '出力※※※※※※※※※※※※※※※※※※※※※※※※※※※
    SplineXYParaFast = OutputArrayXY2D
    
End Function

Private Function SplineParaFast(ByVal ArrayX1D, ByVal ArrayY1D, BunkatuN&, PointCount&)
'パラメトリック関数形式でスプライン補間を行う
'分割して計算を高速化する
'ArrayX1D,ArrayY1Dがどちらも単調増加、単調減少でない場合に用いる。
'20211009

'引数
'ArrayX1D  ・・・補間の対象となるXの値が格納された配列
'ArrayY1D  ・・・補間の対象となるYの値が格納された配列
'BunkatuN  ・・・パラメトリック関数の分割個数（出力されるOutputArrayX1D,OutputArrayY1Dの要素数は(分割個数+1)）
'PointCount・・・分割する際の点数

'返り値
'パラメトリック関数形式で補間されたXList,YList
    
    '入力値のチェック及び修正※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '入力がセルから(ワークシート関数)だった場合の処理
    If IsObject(ArrayX1D) Then
        ArrayX1D = Application.Transpose(ArrayX1D.Value)
    End If
    If IsObject(ArrayY1D) Then
        ArrayY1D = Application.Transpose(ArrayY1D.Value)
    End If
    
    Dim StartNum%
    '行列の開始要素を1に変更（計算しやすいから）
    StartNum = LBound(ArrayX1D, 1) '入力配列の要素の開始番号を取っておく（出力値に合わせるため）
    If LBound(ArrayX1D, 1) <> 1 Then
        ArrayX1D = Application.Transpose(Application.Transpose(ArrayX1D))
    End If
    If LBound(ArrayY1D, 1) <> 1 Then
        ArrayY1D = Application.Transpose(Application.Transpose(ArrayY1D))
    End If
    
    '配列の次元チェック
    Dim JigenCheck1%, JigenCheck2%
    On Error Resume Next
    JigenCheck1 = UBound(ArrayX1D, 2) '配列の次元が1ならエラーとなる
    JigenCheck2 = UBound(ArrayY1D, 2) '配列の次元が1ならエラーとなる
    On Error GoTo 0
    
    '配列の次元が2なら次元1にする。例)配列(1 to N,1 to 1)→配列(1 to N)
    If JigenCheck1 > 0 Then
        ArrayX1D = Application.Transpose(ArrayX1D)
    End If
    If JigenCheck2 > 0 Then
        ArrayY1D = Application.Transpose(ArrayY1D)
    End If
    
    '計算処理※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    N = UBound(ArrayX1D, 1)
    Dim ArrayT1D#(), ArrayParaT1D#()
    
    'X,Yの補間の基準となる配列を作成
    ReDim ArrayT1D(1 To N)
    For I = 1 To N
        '0～1を等間隔
        ArrayT1D(I) = (I - 1) / (N - 1)
    Next I
    
    '出力補間位置の基準位置
    If JigenCheck1 > 0 Then '出力値の形状を入力値に合わせるための処理
        ReDim ArrayParaT1D(StartNum To StartNum - 1 + BunkatuN + 1, 1 To 1)
        For I = 1 To BunkatuN + 1
            '0～1を等間隔
            ArrayParaT1D(StartNum + I - 1, 1) = (I - 1) / (BunkatuN)
        Next I
    Else
        ReDim ArrayParaT1D(StartNum To StartNum - 1 + BunkatuN + 1)
        For I = 1 To BunkatuN + 1
            '0～1を等間隔
            ArrayParaT1D(StartNum + I - 1) = (I - 1) / (BunkatuN)
        Next I
    End If
    
    Dim OutputArrayX1D, OutputArrayY1D
    OutputArrayX1D = SplineByArrayX1DFast(ArrayT1D, ArrayX1D, ArrayParaT1D, PointCount)
    OutputArrayY1D = SplineByArrayX1DFast(ArrayT1D, ArrayY1D, ArrayParaT1D, PointCount)
    
    '出力
    Dim Output(1 To 2)
    Output(1) = OutputArrayX1D
    Output(2) = OutputArrayY1D
    
    SplineParaFast = Output
    
End Function

Private Function SplineByArrayX1DFast(ByVal ArrayX1D, ByVal ArrayY1D, ByVal InputArrayX1D, PointCount&)
 'スプライン補間計算を行う
 '分割して計算することで高速化する

'引数
'HairetuX     ・・・補間の対象となるXの値が格納された配列
'HairetuY     ・・・補間の対象となるYの値が格納された配列
'InputArrayX1D・・・補間位置Xが格納された配列
'PointCount   ・・・分割する際の点数

'返り値
'入力配列InputArrayX1Dに対する補間値の配列
        
    '入力値のチェック及び修正※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '入力がセルから(ワークシート関数)だった場合の処理
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
    '行列の開始要素を1に変更（計算しやすいから）
    If LBound(ArrayX1D, 1) <> 1 Then
        ArrayX1D = Application.Transpose(Application.Transpose(ArrayX1D))
    End If
    If LBound(ArrayY1D, 1) <> 1 Then
        ArrayY1D = Application.Transpose(Application.Transpose(ArrayY1D))
    End If
    StartNum = LBound(InputArrayX1D, 1) 'InputArrayX1Dの開始要素番号を取っておく（出力値を合わせるため）
    If LBound(InputArrayX1D, 1) <> 1 Then
        InputArrayX1D = Application.Transpose(Application.Transpose(InputArrayX1D))
    End If
    
    '配列の次元チェック
    Dim JigenCheck1%, JigenCheck2%, JigenCheck3%
    On Error Resume Next
    JigenCheck1 = UBound(ArrayX1D, 2) '配列の次元が1ならエラーとなる
    JigenCheck2 = UBound(ArrayY1D, 2) '配列の次元が1ならエラーとなる
    JigenCheck3 = UBound(InputArrayX1D, 2) '配列の次元が1ならエラーとなる
    On Error GoTo 0
    
    '配列の次元が2なら次元1にする。例)配列(1 to N,1 to 1)→配列(1 to N)
    If JigenCheck1 > 0 Then
        ArrayX1D = Application.Transpose(ArrayX1D)
    End If
    If JigenCheck2 > 0 Then
        ArrayY1D = Application.Transpose(ArrayY1D)
    End If
    If JigenCheck3 > 0 Then
        InputArrayX1D = Application.Transpose(InputArrayX1D)
    End If

    '計算処理※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Dim SplitArrayList
    SplitArrayList = スプライン補間高速化用に分割処理(ArrayX1D, ArrayY1D, InputArrayX1D, PointCount)
        
    Dim TmpXList, TmpYList, TmpPointList
    Dim Output '出力値格納変数
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

Private Function スプライン補間高速化用に分割処理(ByVal ArrayX1D, ByVal ArrayY1D, ByVal CalPoint1D, PointCount&)
'スプライン補間高速化用に分割処理
'20211009

'引数
'ArrayX1D  ・・・補間元のX座標リスト
'ArrayY1D  ・・・補間元のY座標リスト
'CalPoint1D・・・補間位置のX座標リスト
'PointCount・・・分割後の一つの分割の点数

    Dim I&, J&, II&, JJ&, N&, M&, K&
    N = UBound(ArrayX1D, 1)
    Dim PointN&
    PointN = UBound(CalPoint1D, 1)
    
    Dim Output '出力値格納変数
    ReDim Output(1 To N, 1 To 3) '1:補間元X座標リスト,2:補間元Y座標リスト,3:補間位置X座標リスト
    'Nはとりあえずの最大で、後で配列を縮小する
    
    Dim TmpXList, TmpYList, TmpPointList, TmpInterXList
    Dim StartNum&, EndNum& '分割する補間元座標の開始位置と終了位置
    Dim InterStartNum&, InterEndNum& '分割された補間元座標で実際の補間範囲の開始位置と終了位置
    
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
    
    '出力する格納配列の範囲調整
    Output = ExtractArray(Output, 1, 1, K, 3)
    
    '分割した補間位置で重複するものを消去
    N = UBound(Output, 1)
    Dim TmpList1, TmpList2
    For I = 2 To N
        TmpList1 = Output(I - 1, 3)
        TmpList2 = Output(I, 3)
        If IsEmpty(TmpList1) = False And IsEmpty(TmpList2) = False Then
            If TmpList1(UBound(TmpList1, 1)) = TmpList2(1) Then '最後の要素と最初の要素を比較する
                If UBound(TmpList2, 1) = 1 Then
                    TmpList2 = Empty
                Else
                    TmpList2 = ExtractArray1D(TmpList2, 2, UBound(TmpList2, 1))
                End If
                Output(I, 3) = TmpList2
            End If
        End If
    Next
    
    スプライン補間高速化用に分割処理 = Output
    
End Function

Private Function ExtractByRangeArray1D(InputArray1D, RangeArray1D)
'一次元配列の指定範囲を抽出する。
'指定範囲はRangeArray1Dで指定する。
'20211009

'引数
'InputArray1D・・・抽出元の一次元配列
'RangeArray1D・・・抽出する範囲を指定する一次元配列

'例
'InputArray1D = (1,2,3,4,5,6,7,8,9,10)
'RangeArray1D = (3,4,7)
'出力 = (3,4,5,6,7)

    '引数チェック
    Call CheckArray1D(InputArray1D, "InputArray1D")
    Call CheckArray1DStart1(InputArray1D, "InputArray1D")
    Call CheckArray1D(RangeArray1D, "RangeArray1D")
    Call CheckArray1DStart1(RangeArray1D, "RangeArray1D")
    
    Dim I&, J&, II&, JJ&, N&, M&, K&
    
    
    '指定範囲の最小、最大を取得
    Dim MinNum#, MaxNum#
    MinNum = WorksheetFunction.Min(RangeArray1D)
    MaxNum = WorksheetFunction.Max(RangeArray1D)
    
    '抽出範囲の開始位置、終了位置を計算
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
        '抽出範囲なしでEmptyを返す
        Exit Function
    End If
    
    For I = StartNum To N
        If InputArray1D(I) > MaxNum Then
            EndNum = I - 1
            Exit For
        End If
    Next
    
    If EndNum = 0 Then
        '終了位置が見つからない場合は終了まで全部含む
        EndNum = N
    End If
    
    '範囲抽出
    Dim Output '出力値格納変数
    Output = ExtractArray1D(InputArray1D, StartNum, EndNum)
    
    '出力
    ExtractByRangeArray1D = Output
    
End Function

Private Sub CheckArray1D(InputArray, Optional HairetuName$ = "配列")
'入力配列が1次元配列かどうかチェックする
'20210804

    Dim Dummy%
    On Error Resume Next
    Dummy = UBound(InputArray, 2)
    On Error GoTo 0
    If Dummy <> 0 Then
        MsgBox (HairetuName & "は1次元配列を入力してください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub

Private Sub CheckArray1DStart1(InputArray, Optional HairetuName$ = "配列")
'入力1次元配列の開始番号が1かどうかチェックする
'20210804

    If LBound(InputArray, 1) <> 1 Then
        MsgBox (HairetuName & "の開始要素番号は1にしてください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub

Private Function ExtractArray1D(Array1D, StartNum&, EndNum&)
'一次元配列の指定範囲を配列として抽出する
'20211009

'引数
'Array1D ・・・一次元配列
'StartNum・・・抽出範囲の開始番号
'EndNum  ・・・抽出範囲の終了番号
                                   
    '引数チェック
    Call CheckArray1D(Array1D, "Array1D")
    Call CheckArray1DStart1(Array1D, "Array1D")
    
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    N = UBound(Array1D, 1) '要素数
    
    If StartNum > EndNum Then
        MsgBox ("抽出範囲の開始位置「StartNum」は、終了位置「EndNum」以下でなければなりません")
        Stop
        Exit Function
    ElseIf StartNum < 1 Then
        MsgBox ("抽出範囲の開始位置「StartNum」は1以上の値を入れてください")
        Stop
        Exit Function
    ElseIf EndNum > N Then
        MsgBox ("抽出範囲の終了行「EndNum」は抽出元の一次元配列の要素数" & N & "以下の値を入れてください")
        Stop
        Exit Function
    End If
    
    '処理
    Dim Output
    ReDim Output(1 To EndNum - StartNum + 1)
    
    For I = StartNum To EndNum
        Output(I - StartNum + 1) = Array1D(I)
    Next I
    
    '出力
    ExtractArray1D = Output
    
End Function

Private Function ExtractArray(Array2D, StartRow&, StartCol&, EndRow&, EndCol&)
'二次元配列の指定範囲を配列として抽出する
'20210917

'引数
'Array2D ・・・二次元配列
'StartRow・・・抽出範囲の開始行番号
'StartCol・・・抽出範囲の開始列番号
'EndRow  ・・・抽出範囲の終了行番号
'EndCol  ・・・抽出範囲の終了列番号
                                   
    '引数チェック
    Call CheckArray2D(Array2D, "Array2D")
    Call CheckArray2DStart1(Array2D, "Array2D")
    
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    N = UBound(Array2D, 1) '行数
    M = UBound(Array2D, 2) '列数
    
    If StartRow > EndRow Then
        MsgBox ("抽出範囲の開始行「StartRow」は、終了行「EndRow」以下でなければなりません")
        Stop
        End
    ElseIf StartCol > EndCol Then
        MsgBox ("抽出範囲の開始列「StartCol」は、終了列「EndCol」以下でなければなりません")
        Stop
        End
    ElseIf StartRow < 1 Then
        MsgBox ("抽出範囲の開始行「StartRow」は1以上の値を入れてください")
        Stop
        End
    ElseIf StartCol < 1 Then
        MsgBox ("抽出範囲の開始列「StartCol」は1以上の値を入れてください")
        Stop
        End
    ElseIf EndRow > N Then
        MsgBox ("抽出範囲の終了行「StartRow」は抽出元の二次元配列の行数" & N & "以下の値を入れてください")
        Stop
        End
    ElseIf EndCol > M Then
        MsgBox ("抽出範囲の終了列「StartCol」は抽出元の二次元配列の列数" & M & "以下の値を入れてください")
        Stop
        End
    End If
    
    '処理
    Dim Output
    ReDim Output(1 To EndRow - StartRow + 1, 1 To EndCol - StartCol + 1)
    
    For I = StartRow To EndRow
        For J = StartCol To EndCol
            Output(I - StartRow + 1, J - StartCol + 1) = Array2D(I, J)
        Next J
    Next I
    
    '出力
    ExtractArray = Output
    
End Function

Private Sub CheckArray2D(InputArray, Optional HairetuName$ = "配列")
'入力配列が2次元配列かどうかチェックする
'20210804

    Dim Dummy2%, Dummy3%
    On Error Resume Next
    Dummy2 = UBound(InputArray, 2)
    Dummy3 = UBound(InputArray, 3)
    On Error GoTo 0
    If Dummy2 = 0 Or Dummy3 <> 0 Then
        MsgBox (HairetuName & "は2次元配列を入力してください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub

Private Sub CheckArray2DStart1(InputArray, Optional HairetuName$ = "配列")
'入力2次元配列の開始番号が1かどうかチェックする
'20210804

    If LBound(InputArray, 1) <> 1 Or LBound(InputArray, 2) <> 1 Then
        MsgBox (HairetuName & "の開始要素番号は1にしてください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub

Private Function SplineByArrayX1D(ByVal ArrayX1D, ByVal ArrayY1D, ByVal InputArrayX1D)
    'スプライン補間計算を行う
    '＜出力値の説明＞
    '入力配列InputArrayX1Dに対する補間値の配列YList
    
    '＜入力値の説明＞
    'HairetuX：補間の対象となるXの値が格納された配列
    'HairetuY：補間の対象となるYの値が格納された配列
    'InputArrayX1D:補間位置Xが格納された配列

    '入力値のチェック及び修正※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '入力がセルから(ワークシート関数)だった場合の処理
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
    '行列の開始要素を1に変更（計算しやすいから）
    If LBound(ArrayX1D, 1) <> 1 Then
        ArrayX1D = Application.Transpose(Application.Transpose(ArrayX1D))
    End If
    If LBound(ArrayY1D, 1) <> 1 Then
        ArrayY1D = Application.Transpose(Application.Transpose(ArrayY1D))
    End If
    StartNum = LBound(InputArrayX1D, 1) 'InputArrayX1Dの開始要素番号を取っておく（出力値を合わせるため）
    If LBound(InputArrayX1D, 1) <> 1 Then
        InputArrayX1D = Application.Transpose(Application.Transpose(InputArrayX1D))
    End If
    
    '配列の次元チェック
    Dim JigenCheck1%, JigenCheck2%, JigenCheck3%
    On Error Resume Next
    JigenCheck1 = UBound(ArrayX1D, 2) '配列の次元が1ならエラーとなる
    JigenCheck2 = UBound(ArrayY1D, 2) '配列の次元が1ならエラーとなる
    JigenCheck3 = UBound(InputArrayX1D, 2) '配列の次元が1ならエラーとなる
    On Error GoTo 0
    
    '配列の次元が2なら次元1にする。例)配列(1 to N,1 to 1)→配列(1 to N)
    If JigenCheck1 > 0 Then
        ArrayX1D = Application.Transpose(ArrayX1D)
    End If
    If JigenCheck2 > 0 Then
        ArrayY1D = Application.Transpose(ArrayY1D)
    End If
    If JigenCheck3 > 0 Then
        InputArrayX1D = Application.Transpose(InputArrayX1D)
    End If

    '計算処理※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Dim A, B, C, D
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    
    'スプライン計算用の各係数を計算する。参照渡しでA,B,C,Dに格納
    Dim Dummy
    Dummy = SplineKeisu(ArrayX1D, ArrayY1D)
    A = Dummy(1)
    B = Dummy(2)
    C = Dummy(3)
    D = Dummy(4)
    
    Dim SotoNaraTrue As Boolean
    N = UBound(ArrayX1D, 1) '補間対象の要素数
    
    Dim OutputArrayY1D#() '出力するYの格納
    Dim NX%
    NX = UBound(InputArrayX1D, 1) '補間位置の個数
    ReDim OutputArrayY1D(1 To NX)
    Dim TmpX#, TmpY#
    
    For J = 1 To NX
        TmpX = InputArrayX1D(J)
        SotoNaraTrue = False
        For I = 1 To N - 1
            If ArrayX1D(I) < ArrayX1D(I + 1) Then 'Xが単調増加の場合
                If I = 1 And ArrayX1D(1) > TmpX Then '範囲に入らないとき(開始点より前)
                    TmpY = ArrayY1D(1)
                    SotoNaraTrue = True
                    Exit For
                
                ElseIf I = N - 1 And ArrayX1D(I + 1) <= TmpX Then '範囲に入らないとき(終了点より後)
                    TmpY = ArrayY1D(N)
                    SotoNaraTrue = True
                    Exit For
                    
                ElseIf ArrayX1D(I) <= TmpX And ArrayX1D(I + 1) > TmpX Then '範囲内
                    K = I: Exit For
                
                End If
            Else 'Xが単調減少の場合
            
                If I = 1 And ArrayX1D(1) < TmpX Then '範囲に入らないとき(開始点より前)
                    TmpY = ArrayY1D(1)
                    SotoNaraTrue = True
                    Exit For
                
                ElseIf I = N - 1 And ArrayX1D(I + 1) >= TmpX Then '範囲に入らないとき(終了点より後)
                    TmpY = ArrayY1D(N)
                    SotoNaraTrue = True
                    Exit For
                    
                ElseIf ArrayX1D(I + 1) < TmpX And ArrayX1D(I) >= TmpX Then '範囲内
                    K = I: Exit For
                
                End If
            
            End If
        Next I
        
        If SotoNaraTrue = False Then
            TmpY = A(K) + B(K) * (TmpX - ArrayX1D(K)) + C(K) * (TmpX - ArrayX1D(K)) ^ 2 + D(K) * (TmpX - ArrayX1D(K)) ^ 3
        End If
        
        OutputArrayY1D(J) = TmpY
        
    Next J
    
    '出力※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Dim Output
    
    '出力する配列を入力した配列InputArrayX1Dの形状に合わせる
    If JigenCheck3 = 1 Then '入力のInputArrayX1Dが二次元配列
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
        'ワークシート関数の場合
        SplineByArrayX1D = Application.Transpose(Output)
    Else
        'VBA上での処理の場合
        SplineByArrayX1D = Output
    End If
    
End Function

Private Function SplineKeisu(ByVal ArrayX1D, ByVal ArrayY1D)

    '参考：http://www5d.biglobe.ne.jp/stssk/maze/spline.html
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    Dim A, B, C, D
    N = UBound(ArrayX1D, 1)
    ReDim A(1 To N)
    ReDim B(1 To N)
    ReDim D(1 To N)
    
    Dim h#()
    Dim ArrayL2D#() '左辺の配列 要素数(1 to N,1 to N)
    Dim ArrayR1D#() '右辺の配列 要素数(1 to N,1 to 1)
    Dim ArrayLm2D#() '左辺の配列の逆行列 要素数(1 to N,1 to N)
    
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
    
    '右辺の配列の計算
    For I = 1 To N
        If I = 1 Or I = N Then
            ArrayR1D(I, 1) = 0
        Else
            ArrayR1D(I, 1) = 3 * (ArrayY1D(I + 1) - ArrayY1D(I)) / h(I) - 3 * (ArrayY1D(I) - ArrayY1D(I - 1)) / h(I - 1)
        End If
    Next I
    
    '左辺の配列の計算
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
    
    '左辺の配列の逆行列
    ArrayLm2D = F_Minverse(ArrayL2D)
    
    'Cの配列を求める
    C = F_MMult(ArrayLm2D, ArrayR1D)
    C = Application.Transpose(C)
    
    'Bの配列を求める
    For I = 1 To N - 1
        B(I) = (A(I + 1) - A(I)) / h(I) - h(I) * (C(I + 1) + 2 * C(I)) / 3
    Next I
    
    'Dの配列を求める
    For I = 1 To N - 1
        D(I) = (C(I + 1) - C(I)) / (3 * h(I))
    Next I
    
    '出力
    Dim Output(1 To 4)
    Output(1) = A
    Output(2) = B
    Output(3) = C
    Output(4) = D
    
    SplineKeisu = Output

End Function

Private Function F_MMult(ByVal Matrix1, ByVal Matrix2)
    'F_MMult(Matrix1, Matrix2)
    'F_MMult(配列①,配列②)
    '行列の積を計算
    '20180213改良
    '20210603改良
    
    '入力値のチェックと修正※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '配列の次元チェック
    Dim JigenCheck1%, JigenCheck2%
    On Error Resume Next
    JigenCheck1 = UBound(Matrix1, 2) '配列の次元が1ならエラーとなる
    JigenCheck2 = UBound(Matrix2, 2) '配列の次元が1ならエラーとなる
    On Error GoTo 0
    
    '配列の次元が1なら次元2にする。例)配列(1 to N)→配列(1 to N,1 to 1)
    If IsEmpty(JigenCheck1) Then
        Matrix1 = Application.Transpose(Matrix1)
    End If
    If IsEmpty(JigenCheck2) Then
        Matrix2 = Application.Transpose(Matrix2)
    End If
    
    '行列の開始要素を1に変更（計算しやすいから）
    If UBound(Matrix1, 1) = 0 Or UBound(Matrix1, 2) = 0 Then
        Matrix1 = Application.Transpose(Application.Transpose(Matrix1))
    End If
    If UBound(Matrix2, 1) = 0 Or UBound(Matrix2, 2) = 0 Then
        Matrix2 = Application.Transpose(Application.Transpose(Matrix2))
    End If
    
    '入力値のチェック
    If UBound(Matrix1, 2) <> UBound(Matrix2, 1) Then
        MsgBox ("配列1の列数と配列2の行数が一致しません。" & vbLf & _
               "(出力) = (配列1)(配列2)")
        Stop
        End
    End If
    
    '計算処理※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    Dim M2%
    Dim Output#() '出力する配列
    N = UBound(Matrix1, 1) '配列1の行数
    M = UBound(Matrix1, 2) '配列1の列数
    M2 = UBound(Matrix2, 2) '配列2の列数
    
    ReDim Output(1 To N, 1 To M2)
    
    For I = 1 To N '各行
        For J = 1 To M2 '各列
            For K = 1 To M '(配列1のI行)と(配列2のJ列)を掛け合わせる
                Output(I, J) = Output(I, J) + Matrix1(I, K) * Matrix2(K, J)
            Next K
        Next J
    Next I
    
    '出力※※※※※※※※※※※※※※※※※※※※※※※※※※※
    F_MMult = Output
    
End Function

Private Function F_Minverse(ByVal Matrix)
    '20210603改良
    'F_Minverse(input_M)
    'F_Minverse(配列)
    '余因子行列を用いて逆行列を計算
    
    '入力値チェック及び修正※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '行列の開始要素を1に変更（計算しやすいから）
    If LBound(Matrix, 1) <> 1 Or LBound(Matrix, 2) <> 1 Then
        Matrix = Application.Transpose(Application.Transpose(Matrix))
    End If
    
    '入力値のチェック
    Call 正方行列かチェック(Matrix)
    
    '計算処理※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Dim I%, J%, K%, M%, M2%, N% '数え上げ用(Integer型)
    N = UBound(Matrix, 1)
    Dim Output#()
    ReDim Output(1 To N, 1 To N)
    
    Dim detM# '行列式の値を格納
    detM = F_MDeterm(Matrix) '行列式を求める
    
    Dim Mjyokyo '指定の列・行を除去した配列を格納
    
    For I = 1 To N '各列
        For J = 1 To N '各行
            
            'I列,J行を除去する
            Mjyokyo = F_Mjyokyo(Matrix, J, I)
            
            'I列,J行の余因子を求めて出力する逆行列に格納
            Output(I, J) = F_MDeterm(Mjyokyo) * (-1) ^ (I + J) / detM
    
        Next J
    Next I
    
    '出力※※※※※※※※※※※※※※※※※※※※※※※※※※※
    F_Minverse = Output
    
End Function

Private Sub 正方行列かチェック(Matrix)
    '20210603追加
    
    If UBound(Matrix, 1) <> UBound(Matrix, 2) Then
        MsgBox ("正方行列を入力してください" & vbLf & _
                "入力された配列の要素数は" & "「" & _
                UBound(Matrix, 1) & "×" & UBound(Matrix, 2) & "」" & "です")
        Stop
        End
    End If

End Sub

Private Function F_MDeterm(Matrix)
    '20210603改良
    'F_MDeterm(Matrix)
    'F_MDeterm(配列)
    '行列式を計算
    
    '入力値チェック及び修正※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '行列の開始要素を1に変更（計算しやすいから）
    If LBound(Matrix, 1) <> 1 Or LBound(Matrix, 2) <> 1 Then
        Matrix = Application.Transpose(Application.Transpose(Matrix))
    End If
    
    '入力値のチェック
    Call 正方行列かチェック(Matrix)
    
    '計算処理※※※※※※※※※※※※※※※※※※※※※※※※※※※
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    N = UBound(Matrix, 1)
    
    Dim Matrix2 '掃き出しを行う行列
    Matrix2 = Matrix
    
    For I = 1 To N '各列
        For J = I To N '掃き出し元の行の探索
            If Matrix2(J, I) <> 0 Then
                K = J '掃き出し元の行
                Exit For
            End If
            
            If J = N And Matrix2(J, I) = 0 Then '掃き出し元の値が全て0なら行列式の値は0
                F_MDeterm = 0
                Exit Function
            End If
            
        Next J
        
        If K <> I Then '(I列,I行)以外で掃き出しとなる場合は行を入れ替え
            Matrix2 = F_Mgyoirekae(Matrix2, I, K)
        End If
        
        '掃き出し
        Matrix2 = F_Mgyohakidasi(Matrix2, I, I)
              
    Next I
    
    
    '行列式の計算
    Dim Output#
    Output = 1
    
    For I = 1 To N '各(I列,I行)を掛け合わせていく
        Output = Output * Matrix2(I, I)
    Next I
    
    '出力※※※※※※※※※※※※※※※※※※※※※※※※※※※
    F_MDeterm = Output
    
End Function

Private Function F_Mgyoirekae(Matrix, Row1%, Row2%)
    '20210603改良
    'F_Mgyoirekae(Matrix, Row1, Row2)
    'F_Mgyoirekae(配列,指定行番号①,指定行番号②)
    '行列Matrixの①行と②行を入れ替える
    
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    Dim Output
    
    Output = Matrix
    M = UBound(Matrix, 2) '列数取得
    
    For I = 1 To M
        Output(Row2, I) = Matrix(Row1, I)
        Output(Row1, I) = Matrix(Row2, I)
    Next I
    
    F_Mgyoirekae = Output
End Function

Private Function F_Mgyohakidasi(Matrix, Row%, Col%)
    '20210603改良
    'F_Mgyohakidasi(Matrix, Row, Col)
    'F_Mgyohakidasi(配列,指定行,指定列)
    '行列MatrixのRow行､Col列の値で各行を掃き出す
    
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    Dim Output
    
    Output = Matrix
    N = UBound(Output, 1) '行数取得
    
    Dim Hakidasi '掃き出し元の行
    Dim X# '掃き出し元の値
    Dim Y#
    ReDim Hakidasi(1 To N)
    X = Matrix(Row, Col)
    
    For I = 1 To N '掃き出し元の1行を作成
        Hakidasi(I) = Matrix(Row, I)
    Next I
    
    
    For I = 1 To N '各行
        If I = Row Then
            '掃き出し元の行の場合はそのまま
            For J = 1 To N
                Output(I, J) = Matrix(I, J)
            Next J
        
        Else
            '掃き出し元の行以外の場合は掃き出し
            Y = Matrix(I, Col) '掃き出し基準の列の値
            For J = 1 To N
                Output(I, J) = Matrix(I, J) - Hakidasi(J) * Y / X
            Next J
        End If
    
    Next I
    
    F_Mgyohakidasi = Output
    
End Function

Private Function F_Mjyokyo(Matrix, Row%, Col%)
    '20210603改良
    'F_Mjyokyo(Matrix, Row, Col)
    'F_Mjyokyo(配列,指定行,指定列)
    '行列MatrixのRow行、Col列を除去した行列を返す
    
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    Dim Output '指定した行・列を除去後の配列
    
    N = UBound(Matrix, 1) '行数取得
    M = UBound(Matrix, 2) '列数取得
    ReDim Output(1 To N - 1, 1 To M - 1)
    
    Dim I2%, J2%
    
    I2 = 0 '行方向数え上げ初期化
    For I = 1 To N
        If I = Row Then
            'なにもしない
        Else
            I2 = I2 + 1 '行方向数え上げ
            
            J2 = 0 '列方向数え上げ初期化
            For J = 1 To M
                If J = Col Then
                    'なにもしない
                Else
                    J2 = J2 + 1 '列方向数え上げ
                    Output(I2, J2) = Matrix(I, J)
                End If
            Next J
            
        End If
    Next I
    
    F_Mjyokyo = Output

End Function

Private Function UnionArray1D(UpperArray1D, LowerArray1D)
'一次元配列同士を結合して1つの配列とする。
'20210923

'UpperArray1D・・・上に結合する一次元配列
'LowerArray1D・・・下に結合する一次元配列

    '引数チェック
    Call CheckArray1D(UpperArray1D, "UpperArray1D")
    Call CheckArray1DStart1(UpperArray1D, "UpperArray1D")
    Call CheckArray1D(LowerArray1D, "LowerArray1D")
    Call CheckArray1DStart1(LowerArray1D, "LowerArray1D")
    
    '処理
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
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
    
    '出力
    UnionArray1D = Output
    
End Function

Private Function DrawPolyLine(XYList, TargetSheet As Worksheet) As Shape
'XY座標からポリラインを描く
'シェイプをオブジェクト変数として返す
'20210921

'引数
'XYList         ・・・XY座標が入った二次元配列 X方向→右方向 Y方向→下方向
'TargetSheet    ・・・作図対象のシート

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
'現在カーソル位置のドキュメント座標取得
'カーソル位置のスクリーン座標を、
'カーソルが乗っているセルの四隅のスクリーン座標の関係性から、
'カーソルが乗っているセルの四隅のドキュメント座標をもとに補間して、
'カーソル位置のドキュメント座標を求める。
'20211005

'引数
'[ImmidiateShow]・・・イミディエイトウィンドウに計算結果などを表示するか(デフォルトはTrue)

'返り値
'Output(1 to 2)・・・1:カーソル位置のドキュメント座標X,2:カーソル位置のドキュメント座標Y(Double型)
'カーソルがシート内にない場合はEmptyを返す。

'参考：https://gist.github.com/furyutei/f0668f33d62ccac95d1643f15f19d99a?s=09#to-footnote-1

    Dim Win As Window
    Set Win = ActiveWindow
    
    'カーソルのスクリーン座標取得
    Dim Cursor As PointAPI, CursorScreenX#, CursorScreenY#
    Call GetCursorPos(Cursor)
    CursorScreenX = Cursor.X
    CursorScreenY = Cursor.Y
    
    'カーソルが乗っているセルを取得
    Dim CursorCell As Range, Dummy
    Set Dummy = Win.RangeFromPoint(CursorScreenX, CursorScreenY)
    If TypeName(Dummy) = "Range" Then
        Set CursorCell = Dummy
    Else
        'カーソルがセルに乗ってないので終了
        Exit Function
    End If
    
    '四隅のスクリーン座標を取得
    Dim X1Screen#, X2Screen#, Y1Screen#, Y2Screen# '四隅のスクリーン座標
    Dummy = GetXYCellScreenUpperLeft(CursorCell)
    If IsEmpty(Dummy) Then Exit Function
    X1Screen = Dummy(1)
    Y1Screen = Dummy(2)
    
    Dummy = GetXYCellScreenLowerRight(CursorCell)
    If IsEmpty(Dummy) Then Exit Function
    X2Screen = Dummy(1)
    Y2Screen = Dummy(2)
    
    '四隅のドキュメント座標取得
    Dim X1Document#, X2Document#, Y1Document#, Y2Document# '四隅のドキュメント座標
    X1Document = CursorCell.Left
    X2Document = CursorCell.Left + CursorCell.Width
    Y1Document = CursorCell.Top
    Y2Document = CursorCell.Top + CursorCell.Height
    
    'マウスカーソルのドキュメント座標を補間で計算
    Dim CursorDocumentX#, CursorDocumentY#
    CursorDocumentX = X1Document + (X2Document - X1Document) * (CursorScreenX - X1Screen) / (X2Screen - X1Screen)
    CursorDocumentY = Y1Document + (Y2Document - Y1Document) * (CursorScreenY - Y1Screen) / (Y2Screen - Y1Screen)
        
    '出力
    Dim Output#(1 To 2)
    Output(1) = CursorDocumentX
    Output(2) = CursorDocumentY
    
    GetXYDocumentFromCursor = Output
    
    '確認表示
    If ImmidiateShow Then
        Debug.Print "カーソルの乗ったセル", CursorCell.Address(False, False)
        Debug.Print "カーソルスクリーン座標", "CursorScreenX:" & CursorScreenX, "CursorScreenY:" & CursorScreenY
        Debug.Print "カーソルドキュメント座標", "CursorDocumentX:" & WorksheetFunction.Round(CursorDocumentX, 1), "CursorDocumentY:" & WorksheetFunction.Round(CursorDocumentY, 1)
        Debug.Print "セル左上スクリーン座標", "X1Screen:" & X1Screen, , "Y1Screen:" & Y1Screen
        Debug.Print "セル左上ドキュメント座標", "X1Document:" & X1Document, "Y1Document:" & Y1Document
        Debug.Print "セル右下スクリーン座標", "X2Screen:" & X2Screen, , "Y2Screen:" & Y2Screen
        Debug.Print "セル右下ドキュメント座標", "X2Document:" & X2Document, "Y2Document:" & Y2Document
    End If

End Function

Private Function GetXYCellScreenUpperLeft(TargetCell As Range)
'指定セルの左上のスクリーン座標XYを取得する。
'20211005

'引数
'TargetCell・・・対象のセル(Range型)

'返り値
'Output(1 to 2)・・・1:セル左上のスクリーン座標X,2;セル左上のスクリーン座標Y(Double型)

    'セルが表示されているPane(ウィンドウ枠の固定を考慮した表示エリア)
    Dim Pane As Pane
    Set Pane = GetPaneOfCell(TargetCell)
    If Pane Is Nothing Then Exit Function
       
    '【PointsToScreenPixelsの注意事項】
    '【注】対象セルがシート上で表示されていないと取得不可。一部でも表示されていたら取得可能。
    Dim Output#(1 To 2)
    Output(1) = Pane.PointsToScreenPixelsX(TargetCell.Left)
    Output(2) = Pane.PointsToScreenPixelsY(TargetCell.Top)
    
    GetXYCellScreenUpperLeft = Output
    
End Function

Private Function GetPaneOfCell(TargetCell As Range) As Pane
'指定セルのPaneを取得する
'ウィンドウ枠固定、ウィンドウ分割の設定でも取得できる。
'参考：http://www.asahi-net.or.jp/~ef2o-inue/vba_o/sub05_100_120.html
'20211006

'引数
'TargetCell・・・対象のセル/Range型

'返り値
'指定セルが含まれるPane/Pane型
'指定セルが表示範囲外ならNothing
    
    Dim Win As Window
    Set Win = ActiveWindow
    
    Dim Output As Pane
    Dim I& '数え上げ用(Long型)
    
    ' ウィンドウ分割無しか
    If Not Win.FreezePanes And Not Win.Split Then
        'ウィンドウ枠固定でもウィンドウ分割でもない場合
        ' 表示以外にセルがある場合は無視
        If Intersect(Win.VisibleRange, TargetCell) Is Nothing Then Exit Function
        Set Output = Win.Panes(1)
    Else ' 分割あり
        If Win.FreezePanes Then
            ' ウィンドウ枠固定の場合
            ' どのウィンドウに属するか判定
            For I = 1 To Win.Panes.Count
                If Not Intersect(Win.Panes(I).VisibleRange, TargetCell) Is Nothing Then
                    'Paneの表示範囲に含まれる場合はそのPaneを取得
                    Set Output = Win.Panes(I)
                    Exit For
                End If
            Next I
            
            '見つからなかった場合
            If Output Is Nothing Then Exit Function
        Else
            'ウィンドウ分割の場合
            ' ウィンドウ分割はアクティブペインのみ判定
            If Not Intersect(Win.ActivePane.VisibleRange, TargetCell) Is Nothing Then
                Set Output = Win.ActivePane
            Else
                Exit Function
            End If
        End If
    End If
    
    '出力
    Set GetPaneOfCell = Output
    
End Function

Private Function GetXYCellScreenLowerRight(TargetCell As Range)
'指定セルの右下のスクリーン座標XYを取得する。
'20211005

'引数
'TargetCell・・・対象のセル(Range型)

'返り値
'Output(1 to 2)・・・1:セル右下のスクリーン座標X,2;セル右下のスクリーン座標Y(Double型)

    'セルが表示されているPane(ウィンドウ枠の固定を考慮した表示エリア)
    Dim Pane As Pane
    Set Pane = GetPaneOfCell(TargetCell)
    If Pane Is Nothing Then Exit Function
    
    '【PointsToScreenPixelsの注意事項】
    '【注】対象セルがシート上で表示されていないと取得不可。一部でも表示されていたら取得可能。
    Dim Output#(1 To 2)
    Output(1) = Pane.PointsToScreenPixelsX(TargetCell.Left + TargetCell.Width)
    Output(2) = Pane.PointsToScreenPixelsY(TargetCell.Top + TargetCell.Height)
    
    GetXYCellScreenLowerRight = Output
    
End Function
