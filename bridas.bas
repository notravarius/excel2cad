Attribute VB_Name = "Module1"
Sub DrawBrida()
    Dim AutocadApp As Object
    Dim AutocadDoc As Object
    Dim SectionCoord(0 To 9) As Double
    Dim Rectangle As Object
    Dim Nbrtopbar As Integer
    Dim Nbrbotbar As Integer
    Dim Topsize As Integer
    Dim Botsize As Integer
    Dim Midbar As Integer
    Dim Midsize As Integer
    Dim FilledCir As Integer
    Dim Marray(0) As Object
    Dim OffsetRect As Variant
    Dim Stirrup As Object
    Dim Spacing As Double
    
    Dim bbr1 As Double
   
    Dim bbr2 As Double
    Dim hw As Double
    Dim tw As Double
    Dim bf As Double
    Dim tf As Double

    Dim hci As Double
    Dim hii As Double
    Dim hbi As Double
    Dim vci As Double
    Dim vii As Double
    Dim vbi As Double
    
    Dim hcs As Double
    Dim his As Double
    Dim hbs As Double
    Dim vcs As Double
    Dim vis As Double
    Dim vbs As Double


    Dim hcc As Double
    Dim hic As Double
    Dim hbc As Double
    Dim vc As Double

    Dim diam_agujero As Double
    Dim pos_x As Double


    
    ' //////////////////////////// BRIDA //////////////////////////////////////////////////////

    Dim i As Integer
    
    
    For i = 8 To 22
    
    If IsEmpty(ActiveSheet.Range("B" & i).Value) = False Then
    
     On Error Resume Next
            Set AutocadApp = GetObject(, "Autocad.application")
        On Error GoTo 0
        
        If AutocadApp Is Nothing Then
            Set AutocadApp = CreateObject("Autocad.application")
                AutocadApp.Visible = True
        End If
        
    ' //////////////////////////// VALORES DE LA TABLA //////////////////////////////////////////////////////
    
    bbr1 = ActiveSheet.Range("AS" & i)
    bbr2 = ActiveSheet.Range("AT" & i)
    hw = ActiveSheet.Range("C" & i)
    tw = ActiveSheet.Range("D" & i)
    bf = ActiveSheet.Range("E" & i)
    tf = ActiveSheet.Range("F" & i)
    t_rig = ActiveSheet.Range("AR" & i)

    hci = ActiveSheet.Range("AA" & i)
    hii = ActiveSheet.Range("AB" & i)
    hbi = ActiveSheet.Range("AC" & i)
    
    vci = ActiveSheet.Range("AD" & i)
    vii = ActiveSheet.Range("AE" & i)
    
    vbi = ActiveSheet.Range("AF" & i)
    hcs = ActiveSheet.Range("U" & i)
    his = ActiveSheet.Range("V" & i)
    hbs = ActiveSheet.Range("W" & i)
    vcs = ActiveSheet.Range("X" & i)
    vis = ActiveSheet.Range("Y" & i)
    vc = ActiveSheet.Range("AJ" & i)
    
       

    vbs = ActiveSheet.Range("Z" & i)


    hcc = ActiveSheet.Range("AG" & i)
    hic = ActiveSheet.Range("AH" & i)
    hbc = ActiveSheet.Range("AI" & i)
    
    
    awi = ActiveSheet.Range("AM" & i)
    aws = ActiveSheet.Range("AL" & i)
    aww = ActiveSheet.Range("AN" & i)

    diam_agujero = ActiveSheet.Range("DN" & i)
    pos_x = ActiveSheet.Range("DO" & i)
    dif = ActiveSheet.Range("DP" & i)
    bbr2_adoptado = ActiveSheet.Range("DQ" & i)
    
    nfs = ActiveSheet.Range("O" & i)
    nfi = ActiveSheet.Range("Q" & i)
    nfc = ActiveSheet.Range("S" & i)
    

     ' //////////////////////////////////////////////////////////////////////////////////////////////////////////////////

          
    'Point 1
    SectionCoord(0) = pos_x
    SectionCoord(1) = -dif / 2
    'Point 2
    SectionCoord(2) = SectionCoord(0) + bbr1
    SectionCoord(3) = SectionCoord(1)
    'Point 3
    SectionCoord(4) = SectionCoord(2)
    SectionCoord(5) = bbr2_adoptado - (dif / 2)
    'Point 4
    SectionCoord(6) = SectionCoord(0)
    SectionCoord(7) = SectionCoord(5)
    'Point 1
    SectionCoord(8) = pos_x
    SectionCoord(9) = SectionCoord(1)
    
    On Error Resume Next
        Set AutocadDoc = AutocadApp.activedocument
    On Error GoTo 0
    
    If AutocadDoc Is Nothing Then
        Set AutocadDoc = AutocadApp.document.Add
    End If
    
    AutocadDoc.ActiveLayer = AutocadDoc.Layers("BAU_BRIDA")
     Set Rectangle = AutocadDoc.ModelSpace.addlightweightpolyline(SectionCoord)
         AutocadApp.ZoomExtents
    

    ' /////////////////////////////////////////////////PERFIL////////////////////////////////////////////////////////////////////

    '******************************************************************
    '***********************Ala inferior*******************************
    '******************************************************************
    
    Dim SectionCoordAlaInf(0 To 9) As Double
    Dim RectangleAlaInf As Object
    
    'Point 1
    SectionCoordAlaInf(0) = SectionCoord(0) + bbr1 / 2 - bf / 2
    SectionCoordAlaInf(1) = vci / 2 - tf / 2 + vbi
    'Point 2
    SectionCoordAlaInf(2) = SectionCoordAlaInf(0) + bf
    SectionCoordAlaInf(3) = SectionCoordAlaInf(1)
    'Point 2
    SectionCoordAlaInf(4) = SectionCoordAlaInf(2)
    SectionCoordAlaInf(5) = SectionCoordAlaInf(3) + tf
    'Point 2
    SectionCoordAlaInf(6) = SectionCoordAlaInf(4) - bf
    SectionCoordAlaInf(7) = SectionCoordAlaInf(5)
    'Point 2
    SectionCoordAlaInf(8) = SectionCoordAlaInf(6)
    SectionCoordAlaInf(9) = SectionCoordAlaInf(7) - tf
    
    '******************************************************************
    '***********************Alma***************************************
    '******************************************************************
    
    Dim SectionCoordAlma(0 To 9) As Double
    Dim RectangleAlma As Object
    
    'Point 1
    SectionCoordAlma(0) = SectionCoord(0) + bbr1 / 2 - tw / 2
    SectionCoordAlma(1) = vci / 2 + tf / 2 + vbi
    'Point 2
    SectionCoordAlma(2) = SectionCoordAlma(0) + tw
    SectionCoordAlma(3) = SectionCoordAlma(1)
    'Point 2
    SectionCoordAlma(4) = SectionCoordAlma(2)
    SectionCoordAlma(5) = SectionCoordAlma(3) + hw
    'Point 2
    SectionCoordAlma(6) = SectionCoordAlma(4) - tw
    SectionCoordAlma(7) = SectionCoordAlma(5)
    'Point 2
    SectionCoordAlma(8) = SectionCoordAlma(6)
    SectionCoordAlma(9) = SectionCoordAlma(7) - hw
    
    
    '******************************************************************
    '***********************Ala superior*******************************
    '******************************************************************
    
    Dim SectionCoordAlaSup(0 To 9) As Double
    Dim RectangleAlaSup As Object
    
    'Point 1
    SectionCoordAlaSup(0) = SectionCoord(0) + bbr1 / 2 - bf / 2
    SectionCoordAlaSup(1) = bbr2 - vcs / 2 - vbs - tf / 2
    'Point 2
    SectionCoordAlaSup(2) = SectionCoordAlaSup(0) + bf
    SectionCoordAlaSup(3) = SectionCoordAlaSup(1)
    'Point 2
    SectionCoordAlaSup(4) = SectionCoordAlaSup(2)
    SectionCoordAlaSup(5) = SectionCoordAlaSup(3) + tf
    'Point 2
    SectionCoordAlaSup(6) = SectionCoordAlaSup(4) - bf
    SectionCoordAlaSup(7) = SectionCoordAlaSup(5)
    'Point 2
    SectionCoordAlaSup(8) = SectionCoordAlaSup(6)
    SectionCoordAlaSup(9) = SectionCoordAlaSup(7) - tf
     
    '******************************************************************
    '***********************Rigidizador inferior***********************
    '******************************************************************
    
    Dim SectionCoordRigi_inf(0 To 9) As Double
    Dim RectangleRigi_inf As Object
    
    
    
    'Point 1
    SectionCoordRigi_inf(0) = SectionCoord(0) + bbr1 / 2 - t_rig / 2
    SectionCoordRigi_inf(1) = vbi + vci / 2 - tf / 2
    'Point 2
    SectionCoordRigi_inf(2) = SectionCoordRigi_inf(0) + t_rig
    SectionCoordRigi_inf(3) = SectionCoordRigi_inf(1)
    'Point 2
    SectionCoordRigi_inf(4) = SectionCoordRigi_inf(2)
    SectionCoordRigi_inf(5) = SectionCoordRigi_inf(3) - (bbr2_adoptado - hw - 2 * tf) / 2
    'Point 2
    SectionCoordRigi_inf(6) = SectionCoordRigi_inf(4) - t_rig
    SectionCoordRigi_inf(7) = SectionCoordRigi_inf(5)
    'Point 2
    SectionCoordRigi_inf(8) = SectionCoordRigi_inf(6)
    SectionCoordRigi_inf(9) = SectionCoordRigi_inf(7) + (bbr2_adoptado - hw - 2 * tf) / 2

    '******************************************************************
    '***********************Rigidizador superior***********************
    '******************************************************************
    
    Dim SectionCoordRigi_sup(0 To 9) As Double
    Dim RectangleRigi_sup As Object
    
    
    
    'Point 1
    SectionCoordRigi_sup(0) = SectionCoord(0) + bbr1 / 2 - t_rig / 2
    SectionCoordRigi_sup(1) = bbr2 + (dif / 2)
    'Point 2
    SectionCoordRigi_sup(2) = SectionCoordRigi_sup(0) + t_rig
    SectionCoordRigi_sup(3) = SectionCoordRigi_sup(1)
    'Point 2
    SectionCoordRigi_sup(4) = SectionCoordRigi_sup(2)
    SectionCoordRigi_sup(5) = SectionCoordRigi_sup(3) - (bbr2_adoptado - hw - 2 * tf) / 2
    'Point 2
    SectionCoordRigi_sup(6) = SectionCoordRigi_sup(4) - t_rig
    SectionCoordRigi_sup(7) = SectionCoordRigi_sup(5)
    'Point 2
    SectionCoordRigi_sup(8) = SectionCoordRigi_sup(6)
    SectionCoordRigi_sup(9) = SectionCoordRigi_sup(7) + (bbr2_adoptado - hw - 2 * tf) / 2
    
    '******************************************************************
    '***********************Soldadura**********************************
    '******************************************************************
    
        
    Dim SectionCoordSold(0 To 43) As Double
    Dim RectangleSold As Object
    
    dh_ala_1 = (bf / 2) - (t_rig / 2) - aww
    dh_ala_2 = (bf / 2) - (tw / 2) - aww
    dv_rig_inf = ((bbr2_adoptado - hw - 2 * tf) / 2 - awi)
    dv_rig_sup = ((bbr2_adoptado - hw - 2 * tf) / 2 - aws)
    dv_ala_inf = awi + tf + awi
    dv_ala_sup = aws + tf + aws
    dh_alma = aww + t_rig + aww
    dv_alma = hw - awi - aws
    
    'Point 1
    SectionCoordSold(0) = SectionCoordAlaInf(0)
    SectionCoordSold(1) = SectionCoordAlaInf(1)
        'Point 1
    SectionCoordSold(2) = SectionCoordSold(0)
    SectionCoordSold(3) = SectionCoordSold(1) - awi
    'Point 1
    SectionCoordSold(4) = SectionCoordSold(2) + dh_ala_1
    SectionCoordSold(5) = SectionCoordSold(3)
    'Point 1
    SectionCoordSold(6) = SectionCoordSold(4)
    SectionCoordSold(7) = SectionCoordSold(5) - dv_rig_inf
    'Point 1
    SectionCoordSold(8) = SectionCoordSold(6) + dh_alma
    SectionCoordSold(9) = SectionCoordSold(7)
    'Point 1
    SectionCoordSold(10) = SectionCoordSold(8)
    SectionCoordSold(11) = SectionCoordSold(9) + dv_rig_inf
    'Point 1
    SectionCoordSold(12) = SectionCoordSold(10) + dh_ala_1
    SectionCoordSold(13) = SectionCoordSold(11)
     'Point 1
    SectionCoordSold(14) = SectionCoordSold(12)
    SectionCoordSold(15) = SectionCoordSold(13) + dv_ala_inf
    'Point 1
    SectionCoordSold(16) = SectionCoordSold(14) - dh_ala_2
    SectionCoordSold(17) = SectionCoordSold(15)
    'Point 1
    SectionCoordSold(18) = SectionCoordSold(16)
    SectionCoordSold(19) = SectionCoordSold(17) + dv_alma
    'Point 1
    SectionCoordSold(20) = SectionCoordSold(18) + dh_ala_2
    SectionCoordSold(21) = SectionCoordSold(19)
    'Point 1
    SectionCoordSold(22) = SectionCoordSold(20)
    SectionCoordSold(23) = SectionCoordSold(21) + dv_ala_sup
    'Point 1
    SectionCoordSold(24) = SectionCoordSold(22) - dh_ala_1
    SectionCoordSold(25) = SectionCoordSold(23)
    'Point 1
    SectionCoordSold(26) = SectionCoordSold(24)
    SectionCoordSold(27) = SectionCoordSold(25) + dv_rig_sup
    'Point 1
    SectionCoordSold(28) = SectionCoordSold(26) - dh_alma
    SectionCoordSold(29) = SectionCoordSold(27)
    'Point 1
    SectionCoordSold(30) = SectionCoordSold(28)
    SectionCoordSold(31) = SectionCoordSold(29) - dv_rig_sup
    'Point 1
    SectionCoordSold(32) = SectionCoordSold(30) - dh_ala_1
    SectionCoordSold(33) = SectionCoordSold(31)
    'Point 1
    SectionCoordSold(34) = SectionCoordSold(32)
    SectionCoordSold(35) = SectionCoordSold(33) - dv_ala_sup
    'Point 1
    SectionCoordSold(36) = SectionCoordSold(34) + dh_ala_2
    SectionCoordSold(37) = SectionCoordSold(35)
    'Point 1
    SectionCoordSold(38) = SectionCoordSold(36)
    SectionCoordSold(39) = SectionCoordSold(37) - dv_alma
    'Point 1
    SectionCoordSold(40) = SectionCoordSold(38) - dh_ala_2
    SectionCoordSold(41) = SectionCoordSold(39)
    'Point 1
    SectionCoordSold(42) = SectionCoordSold(40)
    SectionCoordSold(43) = SectionCoordSold(41) - awi - tf
    
    
    
    
    
    On Error Resume Next
        Set AutocadDoc = AutocadApp.activedocument
    On Error GoTo 0
    
    If AutocadDoc Is Nothing Then
        Set AutocadDoc = AutocadApp.document.Add
    End If
    
    Set RectangleAlaInf = AutocadDoc.ModelSpace.addlightweightpolyline(SectionCoordAlaInf)
    Set RectangleAlma = AutocadDoc.ModelSpace.addlightweightpolyline(SectionCoordAlma)
    Set RectangleAlaSup = AutocadDoc.ModelSpace.addlightweightpolyline(SectionCoordAlaSup)
    Set RectangleRigi_inf = AutocadDoc.ModelSpace.addlightweightpolyline(SectionCoordRigi_inf)
    Set RectangleRigi_sup = AutocadDoc.ModelSpace.addlightweightpolyline(SectionCoordRigi_sup)
    
    AutocadDoc.ActiveLayer = AutocadDoc.Layers("BAU_SOLDADURA")
    Set RectangleSold = AutocadDoc.ModelSpace.addlightweightpolyline(SectionCoordSold)
    AutocadDoc.ActiveLayer = AutocadDoc.Layers("BAU_BRIDA")
        AutocadApp.ZoomExtents


   ' /////////////////////////////////////////////////PERNOS O BULONES////////////////////////////////////////////////////////////////////
   Dim CirObj As Object

   '*************************************************************************************************************
    '***********************BULONES INFERIORES********************************************************************
    '*************************************************************************************************************
    
    
    '******************************************************************
    '***********************Bulon 1************************************
    '******************************************************************
    Dim CircleCenter1(0 To 2) As Double
    Dim ColumnDiameter1 As Double
    Dim acadCircle1 As Object
    
    ColumnDiameter1 = diam_agujero / 2
    CircleCenter1(0) = SectionCoord(0) + hbi
    CircleCenter1(1) = vbi
    
    
    '******************************************************************
    '***********************Bulon 2************************************
    '******************************************************************
    Dim CircleCenter2(0 To 2) As Double
    Dim ColumnDiameter2 As Double
    Dim acadCircle2 As Object
    
    ColumnDiameter2 = diam_agujero / 2
    CircleCenter2(0) = SectionCoord(0) + bbr1 - hbi
    CircleCenter2(1) = vbi

    '******************************************************************
    '***********************Bulon 3************************************
    '******************************************************************
      
    Dim CircleCenter3(0 To 2) As Double
    Dim ColumnDiameter3 As Double
    Dim acadCircle3 As Object
    
    ColumnDiameter3 = diam_agujero / 2
    CircleCenter3(0) = SectionCoord(0) + hbi
    CircleCenter3(1) = vbi + vci
    
    '******************************************************************
    '***********************Bulon 4************************************
    '******************************************************************
    Dim CircleCenter4(0 To 2) As Double
    Dim ColumnDiameter4 As Double
    Dim acadCircle4 As Object
    
    ColumnDiameter4 = diam_agujero / 2
    CircleCenter4(0) = SectionCoord(0) + bbr1 - hbi
    CircleCenter4(1) = vbi + vci
    
    '******************************************************************
    '***********************Bulon 5************************************
    '******************************************************************
    Dim CircleCenter5(0 To 2) As Double
    Dim ColumnDiameter5 As Double
    Dim acadCircle5 As Object
    
    ColumnDiameter5 = diam_agujero / 2
    CircleCenter5(0) = SectionCoord(0) + hbi
    CircleCenter5(1) = vbi + vci + vii
    
    '******************************************************************
    '***********************Bulon 6************************************
    '******************************************************************
    Dim CircleCenter6(0 To 2) As Double
    Dim ColumnDiameter6 As Double
    Dim acadCircle6 As Object
    
    ColumnDiameter6 = diam_agujero / 2
    CircleCenter6(0) = SectionCoord(0) + bbr1 - hbi
    CircleCenter6(1) = vbi + vci + vii
    
    If vc <> 0 Then
    
        '******************************************************************
        '***********************Bulon 5 PRIMA******************************
        '******************************************************************
        Dim CircleCenter5prima(0 To 2) As Double
        Dim ColumnDiameter5prima As Double
        Dim acadCircle5prima As Object
        
        ColumnDiameter5prima = diam_agujero / 2
        CircleCenter5prima(0) = SectionCoord(0) + hbi
        CircleCenter5prima(1) = vbi + vci + vii + vc
        
        '******************************************************************
        '***********************Bulon 6 PRIMA******************************
        '******************************************************************
        Dim CircleCenter6prima(0 To 2) As Double
        Dim ColumnDiameter6prima As Double
        Dim acadCircle6prima As Object
        
        ColumnDiameter6prima = diam_agujero / 2
        CircleCenter6prima(0) = SectionCoord(0) + bbr1 - hbi
        CircleCenter6prima(1) = vbi + vci + vii + vc
    
    End If
    
    '*************************************************************************************************************
    '***********************BULONES SUPERIORES********************************************************************
    '*************************************************************************************************************
    
    
    '******************************************************************
    '***********************Bulon 7************************************
    '******************************************************************
    Dim CircleCenter7(0 To 2) As Double
    Dim ColumnDiameter7 As Double
    Dim acadCircle7 As Object
    
    ColumnDiameter7 = diam_agujero / 2
    CircleCenter7(0) = SectionCoord(0) + hbs
    CircleCenter7(1) = bbr2 - vbs
    
    '******************************************************************
    '***********************Bulon 8************************************
    '******************************************************************
    Dim CircleCenter8(0 To 2) As Double
    Dim ColumnDiameter8 As Double
    Dim acadCircle8 As Object
    
    ColumnDiameter8 = diam_agujero / 2
    CircleCenter8(0) = SectionCoord(0) + bbr1 - hbs
    CircleCenter8(1) = bbr2 - vbs
    
    
    '******************************************************************
    '***********************Bulon 9************************************
    '******************************************************************
    Dim CircleCenter9(0 To 2) As Double
    Dim ColumnDiameter9 As Double
    Dim acadCircle9 As Object
    
    ColumnDiameter9 = diam_agujero / 2
    CircleCenter9(0) = SectionCoord(0) + hbs
    CircleCenter9(1) = bbr2 - vbs - vcs

    '******************************************************************
    '***********************Bulon 10************************************
    '******************************************************************
    Dim CircleCenter10(0 To 2) As Double
    Dim ColumnDiameter10 As Double
    Dim acadCircle10 As Object
    
    ColumnDiameter10 = diam_agujero / 2
    CircleCenter10(0) = SectionCoord(0) + bbr1 - hbs
    CircleCenter10(1) = bbr2 - vbs - vcs
    
    '******************************************************************
    '***********************Bulon 11************************************
    '******************************************************************
    Dim CircleCenter11(0 To 2) As Double
    Dim ColumnDiameter11 As Double
    Dim acadCircle11 As Object
    
    ColumnDiameter11 = diam_agujero / 2
    CircleCenter11(0) = SectionCoord(0) + hbs
    CircleCenter11(1) = bbr2 - vbs - vcs - vis
    
    '******************************************************************
    '***********************Bulon 12************************************
    '******************************************************************
    Dim CircleCenter12(0 To 2) As Double
    Dim ColumnDiameter12 As Double
    Dim acadCircle12 As Object
    
    ColumnDiameter12 = diam_agujero / 2
    CircleCenter12(0) = SectionCoord(0) + bbr1 - hbs
    CircleCenter12(1) = bbr2 - vbs - vcs - vis
    
    
    If vc <> 0 Then
    
        '******************************************************************
        '***********************Bulon 11 prima************************************
        '******************************************************************
        Dim CircleCenter11prima(0 To 2) As Double
        Dim ColumnDiameter11prima As Double
        Dim acadCircle11prima As Object
        
        ColumnDiameter11prima = diam_agujero / 2
        CircleCenter11prima(0) = SectionCoord(0) + hbs
        CircleCenter11prima(1) = bbr2 - vbs - vcs - vis - vc
        
        '******************************************************************
        '***********************Bulon 12 prima************************************
        '******************************************************************
        Dim CircleCenter12prima(0 To 2) As Double
        Dim ColumnDiameter12prima As Double
        Dim acadCircle12prima As Object
        
        ColumnDiameter12prima = diam_agujero / 2
        CircleCenter12prima(0) = SectionCoord(0) + bbr1 - hbs
        CircleCenter12prima(1) = bbr2 - vbs - vcs - vis - vc
    
    End If
    
    '******************************************************************
    '***********************Bulon central 1 para nfc=1*******************
    '******************************************************************
    Dim CircleCenterCentral_1(0 To 2) As Double
    Dim ColumnDiameterCentral_1 As Double
    Dim acadCircleCentral_1 As Object
    
    ColumnDiameterCentral_1 = diam_agujero / 2
    CircleCenterCentral_1(0) = SectionCoord(0) + hbs
    CircleCenterCentral_1(1) = bbr2 / 2
    
    '******************************************************************
    '***********************Bulon central 1 para nfc=1*******************
    '******************************************************************
    Dim CircleCenterCentral_2(0 To 2) As Double
    Dim ColumnDiameterCentral_2 As Double
    Dim acadCircleCentral_2 As Object
    
    ColumnDiameterCentral_2 = diam_agujero / 2
    CircleCenterCentral_2(0) = SectionCoord(0) + bbr1 - hbs
    CircleCenterCentral_2(1) = bbr2 / 2
    
    
    
    ' *************************Circle plots***************************************** '
    
    Set acadCircle1 = AutocadDoc.ModelSpace.AddCircle(CircleCenter1, ColumnDiameter1)
    Set acadCircle2 = AutocadDoc.ModelSpace.AddCircle(CircleCenter2, ColumnDiameter2)
    Set acadCircle7 = AutocadDoc.ModelSpace.AddCircle(CircleCenter7, ColumnDiameter7)
    Set acadCircle8 = AutocadDoc.ModelSpace.AddCircle(CircleCenter8, ColumnDiameter8)
    
      
       
    If nfi = 2 Then
        Set acadCircle3 = AutocadDoc.ModelSpace.AddCircle(CircleCenter3, ColumnDiameter3)
        Set acadCircle4 = AutocadDoc.ModelSpace.AddCircle(CircleCenter4, ColumnDiameter4)
    End If
       
    
    If nfi = 3 Then
        Set acadCircle3 = AutocadDoc.ModelSpace.AddCircle(CircleCenter3, ColumnDiameter3)
        Set acadCircle4 = AutocadDoc.ModelSpace.AddCircle(CircleCenter4, ColumnDiameter4)
        Set acadCircle5 = AutocadDoc.ModelSpace.AddCircle(CircleCenter5, ColumnDiameter5)
        Set acadCircle6 = AutocadDoc.ModelSpace.AddCircle(CircleCenter6, ColumnDiameter6)
    End If
    
    If nfc = 2 Then
        Set acadCircle5prima = AutocadDoc.ModelSpace.AddCircle(CircleCenter5prima, ColumnDiameter5prima)
        Set acadCircle6prima = AutocadDoc.ModelSpace.AddCircle(CircleCenter6prima, ColumnDiameter6prima)
        Set acadCircle11prima = AutocadDoc.ModelSpace.AddCircle(CircleCenter11prima, ColumnDiameter11prima)
        Set acadCircle12prima = AutocadDoc.ModelSpace.AddCircle(CircleCenter12prima, ColumnDiameter12prima)
    End If
    
    If nfc = 1 Then
        Set acadCircleCentral_1 = AutocadDoc.ModelSpace.AddCircle(CircleCenterCentral_1, ColumnDiameterCentral_2)
        Set acadCircleCentral_2 = AutocadDoc.ModelSpace.AddCircle(CircleCenterCentral_2, ColumnDiameterCentral_2)
    End If
    
    If nfs = 2 Then
        Set acadCircle9 = AutocadDoc.ModelSpace.AddCircle(CircleCenter9, ColumnDiameter9)
        Set acadCircle10 = AutocadDoc.ModelSpace.AddCircle(CircleCenter10, ColumnDiameter10)
    End If
    
    If nfs = 3 Then
        Set acadCircle11 = AutocadDoc.ModelSpace.AddCircle(CircleCenter11, ColumnDiameter11)
        Set acadCircle12 = AutocadDoc.ModelSpace.AddCircle(CircleCenter12, ColumnDiameter12)
    End If
    
    
       AutocadApp.ZoomExtents


    Set acadCircle = Nothing
    Set acadDoc = Nothing
    Set acadApp = Nothing


    '******************************************************************
    '***********************Leyenda************************************
    '******************************************************************

    Dim acadText As Object
    
    altura_texto = 40
    diam_bulon = "BUL.:  " & ActiveSheet.Range("M" & i)
    nombre_brida = ActiveSheet.Range("A" & i)
    espesor_brida = "Brida:  " & ActiveSheet.Range("AQ" & i) & "mm"
    
         
    Dim TextPosition1(0 To 2) As Double
    TextPosition1(0) = pos_x + 15 'txtpt(0)
    TextPosition1(1) = -70 'txtpt(1)
    TextPosition1(2) = 0 'txtpt(2)
    'MsgBox Val(TextBox10.Text)
    'ModelSpace.AddText "testing", TextPosition, 2
    AutocadDoc.ActiveLayer = AutocadDoc.Layers("BAU_TEXTOS")
    Set acadText = AutocadDoc.ModelSpace.AddText(nombre_brida, TextPosition1, altura_texto)

    
       
    Dim TextPosition2(0 To 2) As Double
    TextPosition2(0) = pos_x + 15 'txtpt(0)
    TextPosition2(1) = -140 'txtpt(1)
    TextPosition2(2) = 0 'txtpt(2)
    'MsgBox Val(TextBox10.Text)
    'ModelSpace.AddText "testing", TextPosition, 2
    Set acadText = AutocadDoc.ModelSpace.AddText(diam_bulon, TextPosition2, altura_texto)
    
    Dim TextPosition3(0 To 2) As Double
    TextPosition3(0) = pos_x + 15 'txtpt(0)
    TextPosition3(1) = -210 'txtpt(1)
    TextPosition3(2) = 0 'txtpt(2)
    'MsgBox Val(TextBox10.Text)
    'ModelSpace.AddText "testing", TextPosition, 2
    Set acadText = AutocadDoc.ModelSpace.AddText(espesor_brida, TextPosition3, altura_texto)
    
    
   
   End If
   
   Next i
    
End Sub

