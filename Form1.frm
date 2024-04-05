VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check7 
      Caption         =   "Check1"
      Height          =   255
      Left            =   2040
      TabIndex        =   14
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Update onward links"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Generate KeyFacts Matrix"
      Height          =   615
      Left            =   3000
      TabIndex        =   12
      Top             =   720
      Width           =   2175
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Check1"
      Height          =   255
      Left            =   2040
      TabIndex        =   11
      Top             =   2640
      Width           =   255
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Check1"
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   2160
      Width           =   255
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Check1"
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SpecCode Build-up"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "KeyFacts assignment"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "KF_type = 1"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "KF_type = 3"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "KF_type = 4"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "KF_type = 2v2"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Check1"
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   1200
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check1"
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   720
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   255
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2760
      Width           =   1140
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim MDB As Database
Dim TBL As Recordset
Dim tbl2 As Recordset

Set MDB = OpenDatabase(App.Path & "\keyfacts.mdb")

'these 2 tables change every month...
Set TBL = MDB.OpenRecordset("species")
Set tbl2 = MDB.OpenRecordset("matrix", dbOpenDynaset)

Dim i As Long
i = 0
Dim X As Long
Dim SSS As Variant

Dim Variable_Length As Long

TBL.MoveFirst
While Not TBL.EOF
    i = i + 1
    X = TBL!SpecCode
    
    
    
    
    tbl2.FindFirst "speccode=" & X
    If tbl2.NoMatch Then
        tbl2.AddNew
        tbl2.Fields("speccode").Value = X
        tbl2.Fields("famcode").Value = TBL!famcode
        tbl2.Update
    End If
    TBL.MoveNext
    
Wend
'MsgBox TBL.RecordCount
'MsgBox i
tbl2.Close
TBL.Close
MDB.Close


End Sub

Private Sub Command2_Click()
Dim MDB As Database
Dim TBL As Recordset

Set MDB = OpenDatabase(App.Path & "\keyfacts.mdb")
Set TBL = MDB.OpenRecordset("matrix", dbOpenDynaset)


Dim i As Long
i = 0
Dim X As Long
Dim SSS As Variant

'start var
Dim tm_for_KF As Long
Dim GetQprime As Recordset
Dim WithMaturity As Recordset
Dim TM01 As Recordset
Dim TM02 As Recordset
Dim TM03 As Recordset
Dim TM04 As Recordset
Dim STR As String


Dim Go4 As Integer

'end var


TBL.MoveFirst
While Not TBL.EOF

    X = TBL!SpecCode
    
    
    '############################################################################
    '####### start main #########################################################
    '############################################################################
    
tm_for_KF = 0

STR = "SELECT Log([k])/Log(10)+2*Log([loo])/Log(10) AS Q, POPGROWTH.* " & _
"From POPGROWTH " & _
"WHERE   (POPGROWTH.SpecCode=" & X & ")   and " & _
"(POPGROWTH.LinfLmax=0)                   and " & _
"(popgrowth.type is not null)             and " & _
"(popgrowth.auxim <> 'Doubtful') " & _
"ORDER BY Log([k])/Log(10)+2*Log([loo])/Log(10);"


Set GetQprime = MDB.OpenRecordset(STR, dbOpenDynaset)

STR = "SELECT SPECIES.SpecCode " & _
"FROM SPECIES INNER JOIN MATURITY ON SPECIES.SpecCode = MATURITY.Speccode " & _
"WHERE (((SPECIES.SpecCode)= " & X & "));"
Set WithMaturity = MDB.OpenRecordset(STR, dbOpenDynaset)



'for key facts
If GetQprime.RecordCount = 0 Then
    Go4 = 0
    If WithMaturity.RecordCount <> 0 Then
    
        
        STR = "SELECT Avg(MATURITY.tm) AS AvgOftm " & _
        "From MATURITY " & _
        "WHERE ( ((MATURITY.Sex)<>'male')                 AND " & _
        "((MATURITY.tm)<>0 And (MATURITY.tm) Is Not Null) AND " & _
        "((MATURITY.Speccode)=" & X & ")); "
        Set TM01 = MDB.OpenRecordset(STR, dbOpenDynaset)
        
        
        If TM01!avgoftm > 0 Then
            tm_for_KF = TM01!avgoftm
        Else
            STR = "SELECT Avg(([agematmin]+[agematmin2])/2) AS vAvg " & _
            "From MATURITY " & _
            "WHERE   (((MATURITY.AgeMatMin)<>0 And (MATURITY.AgeMatMin) Is Not Null) AND " & _
            "((MATURITY.AgeMatMin2)<>0 And (MATURITY.AgeMatMin2) Is Not Null) AND " & _
            "((MATURITY.Sex)<>'male') AND ((MATURITY.Speccode)=" & X & "));"
            Set TM02 = MDB.OpenRecordset(STR, dbOpenDynaset)
            
            If TM02!vAvg > 0 Then
                tm_for_KF = TM02!vAvg
            Else
                STR = "SELECT Avg(MATURITY.agematmin) AS AvgOfagematmin " & _
                "From MATURITY " & _
                "WHERE ( ((MATURITY.Sex)<>'male')                   AND " & _
                "((MATURITY.agematmin)<>0 And (MATURITY.agematmin) Is Not Null)  AND " & _
                "((MATURITY.Speccode)=" & X & "));"
                Set TM03 = MDB.OpenRecordset(STR, dbOpenDynaset)
                
                If TM03!Avgofagematmin > 0 Then
                    tm_for_KF = TM03!Avgofagematmin
                Else
                    STR = "SELECT Avg(POPCHAR.tmax) AS AvgOftmax " & _
                    "From POPCHAR " & _
                    "WHERE ( ((POPCHAR.Sex)<>'male')             AND " & _
                    "((POPCHAR.tmax)<>0 And (POPCHAR.tmax) Is Not Null)  AND " & _
                    "((POPCHAR.Speccode)=" & X & "));"
                    Set TM04 = MDB.OpenRecordset(STR, dbOpenDynaset)
                    If TM04!Avgoftmax > 0 Then
                        tm_for_KF = TM04!Avgoftmax
                        Go4 = 1
                    Else
                        tm_for_KF = 0
                    End If
                End If
            End If
        End If
    Else
        STR = "SELECT Avg(POPCHAR.tmax) AS AvgOftmax " & _
        "From POPCHAR " & _
        "WHERE ( ((POPCHAR.Sex)<>'male')                     AND " & _
        "((POPCHAR.tmax)<>0 And (POPCHAR.tmax) Is Not Null)  AND " & _
        "((POPCHAR.Speccode)=" & X & "));"
        Set TM04 = MDB.OpenRecordset(STR, dbOpenDynaset)
        If TM04!Avgoftmax > 0 Then
            tm_for_KF = TM04!Avgoftmax
            Go4 = 1
        Else
            tm_for_KF = 0
            
        End If
    End If
End If

'for key facts


var_tm = Null
If GetQprime.RecordCount <> 0 Then
    tmp = 1
Else
    If tm_for_KF = 0 Then
        tmp = 2
    Else
        If Go4 = 0 Then
            tmp = 3
            var_tm = tm_for_KF
        Else
            tmp = 4
            var_tm = tm_for_KF
        End If
    End If
End If

    TBL.Edit
    TBL.Fields("kf_type").Value = tmp
    TBL.Fields("tm_for_KF").Value = var_tm
    TBL.Update


GetQprime.Close
WithMaturity.Close
   
    
    
    '############################################################################
    '####### end main ###########################################################
    '############################################################################
    
    TBL.MoveNext
Wend


TM01.Close
TM02.Close
TM03.Close
TM04.Close


'MsgBox TBL.RecordCount
TBL.Close
MDB.Close


End Sub

Private Sub Command3_Click()
'1

Dim MDB As Database
Dim TBL As Recordset

Set MDB = OpenDatabase(App.Path & "\keyfacts.mdb")
Set TBL = MDB.OpenRecordset("SELECT Matrix.*, SPECIES.FamCode, SPECIES.Length, SPECIES.LengthFemale, SPECIES.LTypeMaxM, SPECIES.LTypeMaxF " & _
"FROM Matrix LEFT JOIN SPECIES ON Matrix.SpecCode = SPECIES.SpecCode;", dbOpenDynaset)

Dim i As Long
i = 0
Dim X As Long
Dim SSS As Variant
Dim STR As String



TBL.MoveFirst
While Not TBL.EOF


'start initialize   c3
'TBL.Fields("fecundity_text").Value = fecundity_text
Variable_Length = Null
variable_type1 = Null
variable_infinity = Null
variable_type2 = Null
finalk = Null
variable_q = Null
xto = Null
var_temp = Null
final_mortality = Null
m1st = Null
m2nd = Null
lspan = Null
GENTIME = Null
gtime = Null
vlmaturity = Null
lm_1 = Null
lm_2 = Null
maturity_lt = Null
lmaxyield = Null
lmaxyield_range1 = Null
lmaxyield_range2 = Null
yield_lt = Null
lmaxyield_est = Null
var_a = Null
var_b = Null
finalval = Null
finalval_type = Null
variable_length2 = Null
variable_type3 = Null
nitrogen = Null
protein = Null
np_weight = Null
rg = Null
fecundity_v = Null
fecundity_v1 = Null
fecundity_v2 = Null
vemsy = Null
veopt = Null
vfmsy = Null
vfopt = Null
vLc = Null
Lc_lt = Null
vE = Null
vYR = Null
resiliency = Null
vrm = Null
vlr = Null
lr_lt = Null
mf = Null
tl = Null
finalqb = Null
finalqb_text = Null
vwinf = Null
var_temp_qb = Null
vAfin = Null
'end initialize





    If TBL!kf_type = 1 Then
    'If TBL!Speccode = 2 Or TBL!Speccode = 2 Then
        i = i + 1
        X = TBL!SpecCode
        
    STR = "SELECT STOCKS.Stockcode From stocks " & _
    "where stocks.SpecCode=" & X & "and stocks.level = 'species in general'"
    Set getstockcode = MDB.OpenRecordset(STR, dbOpenDynaset)
    If getstockcode.RecordCount <> 0 Then
        vstockcode = getstockcode!StockCode
    Else
        vstockcode = 0
    End If
        
    '############################################################################
    '####### start main #########################################################
    '############################################################################
    var_temp = Null
        
        
    '############################################################################
    'start lmax for KF_type = 1
    Variable_Length = 1
    If TBL!length <> "" And Not IsNull(TBL!length) Then
        Variable_Length = TBL!length
    Else
        If TBL!lengthfemale <> "" And Not IsNull(TBL!lengthfemale) Then
            Variable_Length = TBL!lengthfemale
        Else
            Variable_Length = 1
        End If
    End If
    If TBL!ltypemaxm <> "" And Not IsNull(TBL!ltypemaxm) Then
        variable_type1 = TBL!ltypemaxm
    Else
        variable_type1 = ""
    End If
    'end lmax for KF_type = 1
    '############################################################################
    
    
    '############################################################################
    'start linf for KF_type = 1
    
    STR = "SELECT Log([k])/Log(10)+2*Log([loo])/Log(10) AS Q, POPGROWTH.* " & _
    "From POPGROWTH " & _
    "WHERE   (POPGROWTH.SpecCode=" & TBL!SpecCode & ")                                                     and " & _
            "(POPGROWTH.LinfLmax=0)                                                      and " & _
            "(popgrowth.type is not null)                                                and " & _
            "(popgrowth.auxim <> 'Doubtful') " & _
    "ORDER BY Log([k])/Log(10)+2*Log([loo])/Log(10)"
    Set GetQprime = MDB.OpenRecordset(STR, dbOpenDynaset)


    xxx = ((GetQprime.RecordCount / 2) + 0.5)
    
    ii = 0
    xxxLoo = 0
    While Not GetQprime.EOF
        ii = ii + 1
        If ii = Round(xxx) Then
            If IsNull(GetQprime!tlinfinity) Then
                variable_infinity = GetQprime!loo
                'variable_q = (log10(k) + 2 * log10(Loo))
                variable_q = ((Log(GetQprime!k) / Log(10)) + 2 * (Log(GetQprime!loo)) / Log(10))
                
                xxxLoo = GetQprime!loo
                xxxVT2 = GetQprime!Type
            Else
                variable_infinity = GetQprime!tlinfinity
                'variable_q = (log10(Val(k)) + 2 * log10(Val(GetQprime!tlinfinity)))
                variable_q = ((Log(GetQprime!k) / Log(10)) + 2 * (Log(GetQprime!tlinfinity) / Log(10)))
                
                xxxLoo = GetQprime!tlinfinity
                xxxVT2 = "TL"
            End If
            
            
'    if #tlinfinity# is "">
'        <!--- 9/23/99 rf modify --->
'        xxxloo = #loo#>
'        xxxVT2 = #type#>
'    else
'        xxxloo = #tlinfinity#>
'        xxxVT2 = "TL">
'    end if
            
            
            
            
            var_temp = GetQprime!Temperature
            'can put a goto to get out of the loop
            
        End If
        GetQprime.MoveNext
    Wend
    
   
   
       
    
    
    'end linf for KF_type = 1
    '############################################################################
    
    
    '############################################################################
    'start K for KF_type = 1
    'variable_q = #numberformat(variable_q,"99.99")#>
    'variable_infinity = #numberformat(variable_infinity,"9999.9")#>
    
    
    'finalk = (10 ^ (variable_q - 2 * log10(variable_infinity)))
    finalk = (10 ^ (variable_q - 2 * (Log(variable_infinity) / Log(10))))
    
    
    'finalk = #numberformat(finalk,"999.99")#>
    'end K for KF_type = 1
    '############################################################################

    
    'start to for KF_type = 1
    xto = -1 * (10 ^ (-0.3922 - 0.2752 * Log(variable_infinity) / Log(10) - 1.038 * Log(finalk) / Log(10)))
    'end to for KF_type = 1
    
    
    
    '############################################################################
    'start var_temp
    
    If IsNull(var_temp) Then
        STR = "SELECT Avg(POPGROWTH.Temperature) AS AvgOfT " & _
        "From POPGROWTH " & _
        "WHERE (((POPGROWTH.SpecCode)=" & TBL!SpecCode & ") AND (Not (POPGROWTH.Temperature)=0 " & _
        "And (POPGROWTH.Temperature) Is Not Null))"
        Set avgt = MDB.OpenRecordset(STR, dbOpenDynaset)
        
        While Not avgt.EOF
            
            var_temp = avgt!avgoft
            avgt.MoveNext
        Wend
    End If
    
    
    If IsNull(var_temp) Then
    STR = "SELECT Avg(([tempmin]+[tempmax])/2) AS vAvg " & _
    "From STOCKS " & _
    "WHERE   (   ((stocks.tempmin)<>0 And (stocks.tempmin) Is Not Null) AND " & _
                "((stocks.tempmax)<>0 And (stocks.tempmax) Is Not Null) AND " & _
                "((stocks.Speccode)=" & TBL!SpecCode & ")    )"
    Set opt_stocks = MDB.OpenRecordset(STR, dbOpenDynaset)
    If opt_stocks!vAvg > 0 Then
        var_temp = opt_stocks!vAvg
    End If
    End If
    
    
    
If IsNull(var_temp) Then

    
    
    
    
    
    
    STR = "SELECT STOCKS.EnvTemp From STOCKS " & _
    "WHERE (((STOCKS.stockcode)=" & vstockcode & "));"
    Set opt_stocks = MDB.OpenRecordset(STR, dbOpenDynaset)
    
    While Not opt_stocks.EOF
        If opt_stocks!envtemp = "boreal" Then
            var_temp = 6
        ElseIf opt_stocks!envtemp = "deep-water" Then
            var_temp = 8
        ElseIf opt_stocks!envtemp = "high altitude" Then
            var_temp = 1
        ElseIf opt_stocks!envtemp = "polar" Then
            var_temp = 1
        ElseIf opt_stocks!envtemp = "subtropical" Then
            var_temp = 17
        ElseIf opt_stocks!envtemp = "temperate" Then
            var_temp = 10
        ElseIf opt_stocks!envtemp = "tropical" Then
            var_temp = 25
        End If
        opt_stocks.MoveNext
    Wend
End If
    
    'end var_temp
    '############################################################################
    
    
    '############################################################################
    'start variable_type2
    
    
    STR = "SELECT POPgrowth.Speccode,popgrowth.temperature From POPgrowth " & _
    "WHERE (((POPgrowth.Speccode)=" & X & "))"
    Set withgrowth = MDB.OpenRecordset(STR, dbOpenDynaset)
    
    If withgrowth.RecordCount = 0 Then
        If variable_type1 <> "" Then
            variable_type2 = variable_type1
        Else
            variable_type2 = ""
        End If
    Else
        variable_type2 = xxxVT2
    End If
    
    If withgrowth.RecordCount = 0 Then
        variable_type2 = variable_type1
    Else
        variable_type2 = variable_type2
    End If
    
    'end variable_type2
    '############################################################################
    
    
    '############################################################################
    'start M and others
    final_mortality = 10 ^ (0.333 - 0.246 * (Log(variable_infinity) / Log(10)) + 0.744 * (Log(finalk) / Log(10)) + 0.01 * var_temp)
    m1st = Null
    m2nd = Null
    If variable_type2 = "TL" Then
        m1st = 10 ^ (0.333 - 0.246 * (Log(variable_infinity) / Log(10)) + 0.744 * (Log(finalk) / Log(10)) + 0.01 * var_temp - 0.18)
        m2nd = 10 ^ (0.333 - 0.246 * (Log(variable_infinity) / Log(10)) + 0.744 * (Log(finalk) / Log(10)) + 0.01 * var_temp + 0.18)
    Else
        'm1st = ""
        'm2nd = ""
    End If
    'end M and others
    '############################################################################
    
    
    '############################################################################
    'start life span
    lspan = (3 / finalk) + xto
    'end life span
    '############################################################################
    
    
    '############################################################################
    'start gentime
    STR = "SELECT popgrowthref,loo,k,type,lm,tlinfinity From POPGROWTH " & _
    "WHERE (((POPGROWTH.SpecCode)=" & X & ")) order by popgrowth.loo,popgrowth.type;"
    Set medgrowth = MDB.OpenRecordset(STR, dbOpenDynaset)
        
If medgrowth.RecordCount = 0 Then
Else
    vlmaturity = 10 ^ (0.898 * (Log(variable_infinity) / Log(10)) - 0.0781)
    
  
    
    
    If (xxxLoo) Then '<!--- meaning with growth, median Qprime --->
        qtest1 = variable_infinity * (3 / (3 + final_mortality / finalk))
        qtest2 = 10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742)
        If qtest1 < qtest2 Then
            varlopt = 10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742)
        Else
            varlopt = variable_infinity * (3 / (3 + final_mortality / finalk))
        End If
    Else
        varlopt = 10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742)
    End If
    
    evlmaturity = 10 ^ (0.898 * (Log(variable_infinity) / Log(10)) - 0.0781)
    
    If evlmaturity >= varlopt Then
        lm100 = evlmaturity + (variable_infinity - evlmaturity) / 4
        GENTIME = xto + (-1 * (Log(1 - lm100 / variable_infinity) / finalk))
    Else
        If Not IsNull(varlopt) Then
            GENTIME = xto + (-1 * (Log(1 - varlopt / variable_infinity) / finalk))
        End If
    End If
End If
        
        
        
    
    'end gentime
    '############################################################################
    
    '############################################################################
    'start tm
    gtime = Round(xto, 2) + (-1 * (Log(1 - Round(vlmaturity, 1) / Round(variable_infinity, 2)) / Round(finalk, 2)))
    
    'If TBL!speccode = 2 Then
    'MsgBox Round(xto, 2)
    'MsgBox Round(vlmaturity, 1)
    'MsgBox Round(variable_infinity, 2)
    'MsgBox Round(finalk, 2)
    'End If
    
    'end tm
    '############################################################################
        
        
        
    '############################################################################
    'start Lm se
    lm_1 = 10 ^ (0.898 * (Log(variable_infinity) / Log(10)) - 0.0781 - 0.127)
    lm_2 = 10 ^ (0.898 * (Log(variable_infinity) / Log(10)) - 0.0781 + 0.127)
    maturity_lt = variable_type2
    'end Lm se
    '############################################################################
    
        
    '############################################################################
    'start lopt
    
    If xxxLoo <> 0 Then '<!--- meaning with growth, median Qprime --->
        qtest1 = variable_infinity * (3 / (3 + final_mortality / finalk))
        qtest2 = 10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742)
        If qtest1 < qtest2 Then
            lmaxyield = 10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742)
            'MsgBox "2"
        Else
            lmaxyield = variable_infinity * (3 / (3 + Round(final_mortality, 2) / Round(finalk, 2)))
            'MsgBox "3"
            'MsgBox lmaxyield
            'MsgBox variable_infinity
            'MsgBox final_mortality
            'MsgBox finalk
        End If
    Else
        lmaxyield = 10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742)
        'MsgBox "4"
    End If
    
    If xxxLoo = 0 Then '<!--- meaning w/o growth, median Qprime --->
        lmaxyield_est = "Estimated from Linf."
    Else
        If qtest1 < qtest2 Then
            lmaxyield_range1 = 10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742 - 0.073)
            lmaxyield_range2 = 10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742 + 0.073)
            lmaxyield_est = "Estimated from Linf."
        Else
            lmaxyield_range1 = Null
            lmaxyield_range2 = Null
            lmaxyield_est = "Estimated from Linf., K and M."
        End If
    End If
        
    yield_lt = variable_type2
    
    'end lopt
    '############################################################################
    
    
    '############################################################################
    'start of l-w
    variable_length2 = Round(variable_infinity, 1)


    'start <!--- get the median of LW --->
    STR = "SELECT  POPLW.SpecCode,POPLW.LengthMin,POPLW.LengthMax,POPLW.Number,POPLW.Sex, POPLW.a, POPLW.b, COUNTREF.paese, " & _
        "poplw.autoctr, poplw.locality,poplw.type, poplw.a , poplw.b FROM COUNTREF INNER JOIN POPLW ON COUNTREF.C_Code = POPLW.C_Code " & _
    "WHERE (((POPLW.SpecCode)=" & TBL!SpecCode & "))        order by poplw.b"
    Set medlwb = MDB.OpenRecordset(STR, dbOpenDynaset)
    
    If medlwb.RecordCount <> 0 Then
        xLW = (medlwb.RecordCount / 2) + 0.5
        xLW = Round(xLW)
        
        'If X = 68 Then
        '    MsgBox xLW
        'End If
        
        ii = 0
        While Not medlwb.EOF
            ii = ii + 1
            If ii = xLW Then
                v_a = medlwb!a
                v_b = medlwb!b
                lwtype = medlwb!Type
            End If
            medlwb.MoveNext
        Wend
    End If
    'end <!--- get the median of LW --->



'<!--- start of lw --->
If medlwb.RecordCount <> 0 Then
    'start parang hindi dada-anan ito
    'if #parameterexists(variable_length2)# is "No">
    '    variable_length2 = #numberformat(length,"9999.9")#>
    'end if
    'end parang hindi dada-anan ito
    
    
    'if #parameterexists(var_a)# is "No">
    var_a = v_a
    'end if
    'if #parameterexists(var_b)# is "No">
    var_b = v_b
    'end if
    finalval = ((variable_length2 ^ Round(var_b, 2)) * Round(var_a, 4))
    'finalval = ((46 ^ 3.13) * 0.0054)
    
    'If X = 68 Then
    '    MsgBox variable_length2 & " " & var_b & " " & var_a & " = " & finalval
    'End If
    
    variable_type3 = lwtype
    eli01 = Len(Trim(Round(finalval, 1))) * 2


    'If TBL!speccode = 2 Then
    '    MsgBox var_a
    '    MsgBox var_b
    '    MsgBox finalval
    'End If

    
    np_weight = finalval
    
    If finalval > 200 Then
        If finalval > 20000 Then
            finalval = Round(finalval / 1000, 1)
            finalval_type = "kg"
        Else
            finalval = Round(finalval)
            finalval_type = "g"
        End If
    Else
        finalval = Round(finalval, 1)
        finalval_type = "g"
    End If
    
       
    'W = finalval
    'a = Var_a
    'b = Var_b
   
    'Nitrogen & protein:
    
    '<!--- start new weight --->
    If finalval <> 0 Then
        'NP_weight = finalval   .... transferred up to preserve the decimal places..
        
        'If Trim(finalval_type) = "kg" Then
        '    NP_weight = NP_weight * 1000
        'End If
        
        
        
        nitrogen = 10 ^ (1.03 * (Log(np_weight) / Log(10)) - 1.65)
        nitrogen = Round(nitrogen, 1)
        protein = Round(6.25 * nitrogen, 1)
        np_weight = Round(np_weight)
        
        
        
    Else
        np_weight = Null
        nitrogen = Null
        protein = Null
    End If
    '<!--- end new weight --->
    
Else
    '<!--- don't show --->
End If

'<!--- end of lw --->
    
    
    'end of l-w
    '############################################################################
    
    
    '############################################################################
    'start reproductive_guild
    rg = ""
    If withgrowth.RecordCount <> 0 Then
    
    'MsgBox withgrowth.RecordCount
    'MsgBox "riz = " & vstockcode
    
    If withgrowth.RecordCount <> 0 Then
    
    
    
    STR = "SELECT REPRODUC.RepGuild1, REPRODUC.RepGuild2 From REPRODUC " & _
    "WHERE (((REPRODUC.StockCode)=" & vstockcode & "))"
    Set getguild = MDB.OpenRecordset(STR, dbOpenDynaset)
    rg = ""
    If getguild.RecordCount <> 0 Then
        If Trim(getguild!repguild1) = "" Then
            rg = " "
        Else
            rg = getguild!repguild1
        End If
        rg = rg + ": "
        If Trim(getguild.repguild2) = "" Then
            rg = rg & " "
        Else
            rg = rg & getguild!repguild2
        End If
        rg = Trim(rg)
    End If
    End If
    End If
    'end reproductive_guild
    '############################################################################
    

    
    
    
    '############################################################################
    'start fecundity
    
    fecundity_v = Null
    fecundity_v1 = Null
    fecundity_v2 = Null
    fecundity_text = Null
    
    
    
    If withgrowth.RecordCount <> 0 And vstockcode <> 0 Then
    
    'If X = 69 Then
    'MsgBox X
    'MsgBox Vstockcode
    'End If

STR = "SELECT Min(SPAWNING.FecundityMin) AS MinOfFecundityMin From SPAWNING " & _
"GROUP BY SPAWNING.StockCode, SPAWNING.SpecCode " & _
"HAVING ((Not (Min(SPAWNING.FecundityMin))=0) AND ((SPAWNING.StockCode)=" & vstockcode & ") AND ((SPAWNING.SpecCode)=" & X & ")) " & _
"ORDER BY Min(SPAWNING.FecundityMin)"
Set getmin = MDB.OpenRecordset(STR, dbOpenDynaset)

STR = "SELECT Max(SPAWNING.FecundityMax) AS MaxOfFecundityMax From SPAWNING GROUP BY SPAWNING.StockCode, SPAWNING.SpecCode " & _
"HAVING (((SPAWNING.StockCode)=" & vstockcode & ") AND ((SPAWNING.SpecCode)=" & X & ") AND (Not (Max(SPAWNING.FecundityMax))=0))"
Set getmax = MDB.OpenRecordset(STR, dbOpenDynaset)
    
    
If getmin.RecordCount <> 0 Or getmax.RecordCount <> 0 Then
    
    y01 = 0
    y02 = 0
    fecundity_v1 = ""
    fecundity_v2 = ""
    
    
    If getmin.RecordCount <> 0 And getmax.RecordCount <> 0 Then
    If getmin!minoffecunditymin > 0 And getmax!maxoffecunditymax > 0 Then
    y01 = Log(getmin.minoffecunditymin) / Log(10)
    y02 = Log(getmax.maxoffecunditymax) / Log(10)
    y01 = y01 + y02
    y01 = y01 / 2
    fecundity_v = Round(10 ^ y01)
    End If
    End If
    
    If getmin.RecordCount <> 0 Then
        If getmin!minoffecunditymin = "" Then
            fecundity_v1 = "no value (min.)"
        Else
            fecundity_v1 = Round(getmin!minoffecunditymin)
        End If
    Else
        fecundity_v1 = "no record (min.)"
    End If
    
    
    If getmax.RecordCount <> 0 Then
        If getmax!maxoffecunditymax = "" Then
            fecundity_v2 = "no value (max.)"
        Else
            fecundity_v2 = Round(getmax!maxoffecunditymax)
        End If
    Else
        fecundity_v2 = "no record (max.)"
    End If
    
    
    
    If getmin.RecordCount <> 0 And getmax.RecordCount <> 0 Then
        fecundity_text = "Estimated as geometric mean."
    End If


End If
    End If
    'end fecundity
    '############################################################################
    
    
    
'############################################################################
'start yrecruit

'<!--- start yrecruit --->

    vemsy = Null
    veopt = Null
    vfmsy = Null
    vfmsy_rm = Null
    vfopt = Null
    vLc = Null
    Lc_lt = Null
    vE = Null
    vYR = Null





vE = 0.5
vLc = Round(0.4 * variable_infinity, 1)
If final_mortality > 0 Then
    
    vU = 1 - vLc / Round(variable_infinity, 2)
    MK = Round(final_mortality, 2) / Round(finalk, 2)
    
    vx1 = 0
    vx2 = 0
    firstloop = "Y"
    
    oldy = 0
    vlope = 0
    vemsy = 0
    veopt = 0
    vfmsy = 0
    vfmsy_rm = 0
    vfopt = 0
    
    
    While vx1 <= 1
        
        vx1 = vx2
        vx2 = vx2 + 0.001

        vm1 = (1 - vx1) / MK
        vy1 = vx1 * (vU ^ MK) * (1 - ((3 * vU) / (1 + vm1)) + ((3 * vU ^ 2) / (1 + 2 * vm1)) - ((vU ^ 3) / (1 + 3 * vm1)))
        vm2 = (1 - vx2) / MK
        vy2 = vx2 * (vU ^ MK) * (1 - ((3 * vU) / (1 + vm2)) + ((3 * vU ^ 2) / (1 + 2 * vm2)) - ((vU ^ 3) / (1 + 3 * vm2)))
        vslope = (vy2 - vy1) / (vx2 - vx1)
        
        If oldy <> 0 Then
            If vy1 >= oldy Then
            Else
                vemsy = Round(vx1 - 0.001, 2)
                '<cfbreak>
                GoTo EliGo
            End If
        End If
        
        
        oldy = vy1
        
        If firstloop = "Y" Then
            firstvalue = (vy2 - vy1) / (vx2 - vx1)
            firstloop = "N"
        End If

        
    
        If veopt = 0 Then
        If Round(vslope, 3) = Round(firstvalue / 10, 3) Then
            veopt = vx1
        End If
        End If
        
        
    Wend
EliGo:
    '<!--- end get e  --->
    

    vm = (1 - vE) / MK
    lijosh = vE * (vU ^ MK) * (1 - ((3 * vU) / (1 + vm)) + ((3 * vU ^ 2) / (1 + 2 * vm)) - ((vU ^ 3) / (1 + 3 * vm)))
    '<input type="text" name="vYR" value=round(lijosh,'9.9999')#" size="6" onFocus="noedit(2)"  align="right">
    vYR = Round(lijosh, 4)

Else
    '<input type="text" name="vYR" value="" size="6" onFocus="noedit(2)"  align="right">
    vYR = Null
End If

    vLc = Round(vLc, 1)
    Lc_lt = Trim(variable_type2)
    e = vE = Round(vE, 2)
    
    If final_mortality > 0 Then
    
        veopt = Round(veopt, 2)
        vemsy = Round(vemsy, 2)
        vfmsy = Round(final_mortality * vemsy / (1 - vemsy), 2)
        vfmsy_rm = final_mortality * vemsy / (1 - vemsy)
        vfopt = Round(final_mortality * veopt / (1 - veopt), 2)
        
        veopt = Round(veopt, 2)
        vemsy = Round(vemsy, 2)
        vfmsy = Round(vfmsy, 2)
        vfopt = Round(vfopt, 2)
    Else
        vemsy = Null
        veopt = Null
        vfmsy = Null
        vfmsy_rm = Null
        vfopt = Null
    End If





'<!--- end yrecruit   width  msy --->


'end yrecruit
'############################################################################
    
    
    '############################################################################
    'start resiliency
    resiliency = Null
    If Round(finalk, 2) <= 0.05 Then
        resiliency = "Very low; decline threshold 0.70"
    ElseIf Round(finalk, 2) <= 0.15 Then
        resiliency = "Low; decline threshold 0.85"
    ElseIf Round(finalk, 2) <= 0.3 Then
        resiliency = "Medium; decline threshold 0.95"
    ElseIf Round(finalk, 2) > 0.3 Then
        resiliency = "High; decline threshold 0.99"
    Else
        resiliency = "Please enter values for K."
    End If
    'end resiliency
    '############################################################################
    
    
    
    
    
    
    '############################################################################
    'start rm
    vlr = Round(0.4 * variable_infinity, 1)
    vfmsy = Round(vfmsy, 2)
    vrm = Round(2 * vfmsy, 2)
    vypdt = Log(2) / vrm
    lr_lt = Trim(variable_type2)
    
    'If X = 2 Then
    '    MsgBox vfmsy
    '    MsgBox vrm
    'End If
    
    'end rm
    '############################################################################
    
    
    
    '############################################################################
    'start eco
    STR = "SELECT  ECOLOGY.StockCode,ECOLOGY.DietTroph,ECOLOGY.DietSeTroph,ECOLOGY.FoodTroph, " & _
    "ECOLOGY.FoodSeTroph,0 as EcoTroph,0 as EcoSeTroph, '' as mainfood,ECOLOGY.herbivory2 " & _
    "From ECOLOGY WHERE (((ECOLOGY.StockCode)=" & vstockcode & "))"
    Set witheco = MDB.OpenRecordset(STR, dbOpenDynaset)
    
mf = Null
tl = Null
If witheco.RecordCount <> 0 Then
    If witheco!mainfood <> "" Then
        mf = witheco!mainfood
    End If
    If witheco!herbivory2 <> "" Then
        mf = mf & " " & witheco!herbivory2
    End If
    
    
    'Trophic level:
    tl = Null
    If witheco!dietTroph <> "" Then
        tl = Round(witheco!dietTroph, 1)
        If witheco!DietSeTroph <> 0 Then
            tl = tl & "&nbsp;&nbsp;&nbsp;"
            tl = tl & "+/- s.e. " & Round(witheco!DietSeTroph, 2)
        End If
        tl = tl & " Estimated from diet data."
    Else
        If witheco!foodTroph <> "" Then
            tl = tl & " " & Round(witheco!foodTroph, 1)
            If witheco!FoodSeTroph <> 0 Then
                tl = tl & "+/- s.e. " & Round(witheco!FoodSeTroph, 2)
            End If
            tl = tl & " Estimated from food data."
        Else
            If witheco!EcoTroph <> "" Then
                tl = tl & " " & Round(witheco!EcoTroph, 1)
                If witheco!EcoSeTroph <> 0 Then
                    tl = tl & "+/- s.e. " & Round(witheco!EcoSeTroph, 2)
                End If
                tl = tl & " Estimated from Ecopath model."
            End If
        End If
    End If
End If


'<!---start of food consumption  </cfif --->

    
    'end eco
    '############################################################################
    
    
    
    
    
    
    
    
    '############################################################################
    'start qb
    
    finalqb = Null
    finalqb_text = Null
    vwinf = Null
    var_temp_qb = Null
    vAfin = Null
    
    
    
    
    
'<!--- start food consumption --->
contqb = "Y"
'<!--- start get median popqb --->
STR = "SELECT popqb.popqb FROM popqb WHERE (((popqb.speccode)=" & X & ")) ORDER BY popqb.popqb"
Set qbmedian = MDB.OpenRecordset(STR, dbOpenDynaset)


'start of old
'vmedian = (qbmedian.RecordCount / 2) + 0.5
'vmedian = Round(vmedian)
'end of old


vMedian = (qbmedian.RecordCount / 2) + 0.5
int_part = Int(vMedian)
str_conv = "" & vMedian

If Len(str_conv) = 1 Then
    vMedian = Round(vMedian)
Else

    If Mid(str_conv, Len(str_conv) - 1, 2) = ".5" Then
        vMedian = int_part + 1
    Else
        vMedian = Round(vMedian)
    End If
End If




vpopqb = 0


'If X = 69 Then
'    MsgBox qbmedian.RecordCount
'    MsgBox (qbmedian.RecordCount / 2) + 0.5
'    MsgBox Int((qbmedian.RecordCount / 2) + 0.5)
'    MsgBox Round(vmedian)
'    MsgBox Round(vmedian, 1)
'End If


ii = 0
While Not qbmedian.EOF

    ii = ii + 1
    If ii = vMedian Then
        vpopqb = qbmedian!popqb
    End If
    qbmedian.MoveNext
Wend


'<!--- end get median popqb --->

If vpopqb > 0 Then
    explain = "with popqb record"
    contqb = "N"
Else
    explain = "no popqb record"
    
    '<!---#############################################################################################--->
    '<!---### start of no popqb #######################################################################--->
    '<!---#############################################################################################--->
    If medlwb.RecordCount# > 0 Then
        explain = explain & "; with lw rel"
    Else
        explain = explain & "; no lw rel"
    End If

    '<!--- start of A --->
    STR = "SELECT Swimming.AspectRatio as aspectratio From Swimming WHERE (((Swimming.SpecCode)=" & X & "));"
    Set getAR = MDB.OpenRecordset(STR, dbOpenDynaset)
    vAfin = 0
    If getAR.RecordCount > 0 Then
       While Not getAR.EOF
            If getAR!aspectratio > 0 Then
                vAfin = getAR!aspectratio
                explain = explain & "; with aspect ratio"
            Else
                vAfin = 0
                explain = explain & "; w/o aspect ratio"
            End If
            getAR.MoveNext
        Wend
    Else
        explain = explain & "; w/o aspect ratio"
    End If
    '<!--- end of A --->
    
    
    '<!--- start of h and d --->
    SelectHD = "N"
    '<!---
    '    go ecology feeding type
    '    if none troph diettroph
    '        if none troph foodtroph items
    '            if none troph ecotroph
    '--->


    STR = "SELECT ECOLOGY.Herbivory2, 0 as EcoTroph, ECOLOGY.DietTroph, ECOLOGY.FoodTroph From ECOLOGY " & _
    "WHERE (((ECOLOGY.StockCode)=" & vstockcode & "))"
    Set getft = MDB.OpenRecordset(STR, dbOpenDynaset)

    
    
    If getft.RecordCount <> 0 Then
    If Trim(getft!herbivory2) <> "" Then
    
        If Trim(getft!herbivory2) = "mainly animals (troph. 2.8 and up)" Then
            vh = 0
            vd = 0
        ElseIf Trim(getft!herbivory2) = "mainly animals (troph. 2.8 up)" Then
            vh = 0
            vd = 0
        ElseIf Trim(getft!herbivory2) = "mainly plants/detritus (troph. 2-2.19)" Then
            vh = 0
            vd = 1
        ElseIf Trim(getft.herbivory2) = "plants/detritus+animals (troph. 2.2-2.79)" Then
            vh = 1
            vd = 0
        End If
        explain = explain & "; w/ feeding type"
        WithFT = 1
    
    Else
    
    
        explain = explain & "; w/o feeding type"
        WithFT = 0
        If (getft!dietTroph) > 0 Then
            If getft!dietTroph >= 2 And getft!dietTroph <= 2.19 Then
                vh = 0
                vd = 1
            ElseIf getft!dietTroph >= 2.2 And getft!dietTroph <= 2.79 Then
                vh = 1
                vd = 0
            ElseIf getft!dietTroph >= 2.8 Then
                vh = 0
                vd = 0
            End If
            explain = explain & "; from diettroph"
        Else
            If getft!foodTroph > 0 Then
                If getft!foodTroph >= 2 And getft.foodTroph <= 2.19 Then
                    vh = 0
                    vd = 1
                ElseIf getft.foodTroph >= 2.2 And getft.foodTroph <= 2.79 Then
                    vh = 1
                    vd = 0
                ElseIf getft!foodTroph >= 2.8 Then
                    vh = 0
                    vd = 0
                End If
                explain = explain & "; from foodtroph"
            Else
                If getft!EcoTroph > 0 Then
                    If getft!EcoTroph >= 2 And getft!EcoTroph <= 2.19 Then
                        vh = 0
                        vd = 1
                    ElseIf getft!EcoTroph >= 2.2 And getft!EcoTroph <= 2.79 Then
                        vh = 1
                        vd = 0
                    ElseIf getft!EcoTroph >= 2.8 Then
                        vh = 0
                        vd = 0
                    End If
                    explain = explain & "; from ecotroph"
                Else
                    explain = explain & "; no diet,food,eco trophs; select h,d"
                    '<!--- blank means yes
                    'contqb = "Y">
                    '--->
                    SelectHD = "Y"
                    '<!---
                    'contqb = "N">
                    'cont. with the search
                    'now the genus of median of diet,food,eco
                    '--->
                End If
            End If
        End If
    End If
    End If
    '<!--- end of h and d --->

'<!---#############################################################################################--->
'<!---### end of no popqb #########################################################################--->
'<!---#############################################################################################--->
End If
'<!--- end food consumption --->









'<!---start of food consumption  </cfif --->

'<td align="left">Food consumption (Q/B):</td>
If vpopqb > 0 Then
    finalqb = Round(vpopqb, 2)
    finalqb_text = "times the body weight per year"
Else
    '<!---#############################################################################################--->
    '<!---#############################################################################################--->
    '<!---#############################################################################################--->
    If contqb = "Y" Then
        '<!--- start for winf --->
        If medlwb.RecordCount > 0 Then
            'orig
            'vwinf = finalval
            
            'use this instead as the finalval is sometimes in kg
            vwinf = np_weight
            'If X = 68 Then
            '    MsgBox "111"
            'End If
            
            whereWinf = "lw"
        Else
            vwinf = 0.01 * variable_infinity ^ 3
            'If X = 68 Then
            '    MsgBox "222"
            'End If
            
            whereWinf = "DP"
        End If
        '<!--- end for winf --->
        If SelectHD = "Y" Then
            vh = 3
            vd = 3
        End If
        'start vb comment
        '<input type="hidden" name="vh" value="#vh#" size="1" onFocus="noedit(2)"  >
        '<input type="hidden" name="vd" value="#vd#" size="1" onFocus="noedit(2)"  >
        'end vb comment
        
        If vAfin > 0 Then
            xyz = vAfin
        Else
            xyz = 1.32
        End If
        If vh = 3 Then
            '<!---w/o h, d --->
            elix = 10 ^ (7.964 - 0.204 * Log(vwinf) / Log(10) - 1.965 * (1000 / (var_temp + 273.15)) + 0.083 * xyz + 0.532 * 1 + 0.398 * 0)
            elix2 = 10 ^ (7.964 - 0.204 * Log(vwinf) / Log(10) - 1.965 * (1000 / (var_temp + 273.15)) + 0.083 * xyz + 0.532 * 0 + 0.398 * 0)
            elix = (elix2 + elix) / 2
        Else
            '<!---with h, d--->
            If vwinf <> 0 Then
            elix = 10 ^ (7.964 - 0.204 * Log(vwinf) / Log(10) - 1.965 * (1000 / (var_temp + 273.15)) + 0.083 * xyz + 0.532 * vh + 0.398 * vd)
            End If
        End If
        
        
        If whereWinf = "none" Then
            finalqb = Null
        Else
            finalqb = Round(elix, 1)
        End If
        finalqb_text = "times the body weight per year"
        
        
        'Enter Winf, temperature, aspect ratio (A), and food type to estimate Q/B
        If whereWinf = "none" Then
            'Winf =
            vwinf = Null
            'If X = 68 Then
            '    MsgBox "333"
            'End If
            
            'Temp. =
            var_temp_qb = Round(var_temp, 1)
        Else
            'Winf =
            vwinf = Round(vwinf, 1)
            
            'If X = 68 Then
            '    MsgBox "444"
            'End If
            
            'Temp. =
            var_temp_qb = Round(var_temp, 1)
        End If
        
        If vAfin > 0 Then
            'A =
            vAfin = Round(vAfin, 2)
        Else
            vAfin = 1.32
            'A =
            vAfin = Round(vAfin, 2)
            
            '<!---eeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeee--->
            
            'start vb comment
            '<td colspan="8"><img src="../jpgs/Tails.gif" height=29 border=0 alt=""></td>
            '<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
            '<td><input type="radio" name="eli"          onClick="finc('#contqb#',6.55)"></td>
            '<td><input type="radio" name="eli"          onClick="finc('#contqb#',1.9)"></td>
            '<td><input type="radio"  checked name="eli" onClick="finc('#contqb#',1.32)"></td>
            '<td><input type="radio" name="eli"          onClick="finc('#contqb#',1.32)"></td>
            '<td><input type="radio" name="eli"          onClick="finc('#contqb#',1.63)"></td>
            'end vb comment
            
        End If
        If SelectHD = "Y" Then
            'start vb comment
            '<input type="hidden" name="omni" value="1">
            '<td align="center"><font size="2">Detrivore</font></td>
            '<td align="center"><font size="2">Herbivore</font></td>
            '<td align="center"><font size="2">Omnivore</font></td>
            '<td align="center"><font size="2">Carnivore</font></td>
            '<td align="center"><input type="radio" name="riz"           onClick="changeHD('#contqb#',0,1)"></td>
            '<td align="center"><input type="radio" name="riz"          onClick="changeHD('#contqb#',1,0)"></td>
            '<td align="center"><input type="radio" name="riz" checked  onClick="changeHD('#contqb#',3,3)"></td>
            '<td align="center"><input type="radio" name="riz"           onClick="changeHD('#contqb#',0,0)"></td>
            '</tr></table>
            'end vb comment
            
        Else
            'start vb comment
            '<input type="hidden" name="omni" value="0">
            '<td align="center"><font size="2">Detrivore</font></td>
            '<td align="center"><font size="2">Herbivore</font></td>
            '<td align="center"><font size="2">Omnivore</font></td>
            '<td align="center"><font size="2">Carnivore</font></td>
            '<td align="center"><input type="radio" name="riz" if #vh# is 0 and #vd# is 1>checkedend if onClick="changeHD('#contqb#',0,1)"></td>
            '<td align="center"><input type="radio" name="riz" if #vh# is 1 and #vd# is 0>checkedend if onClick="changeHD('#contqb#',1,0)"></td>
            '<td align="center"><input type="radio" name="riz"                                              onClick="changeHD('#contqb#',3,3)"></td>
            '<td align="center"><input type="radio" name="riz" if #vh# is 0 and #vd# is 0>checkedend if onClick="changeHD('#contqb#',0,0)"></td>
            'end vb comment
        End If
    End If '<!--- if #contqb# is "Y"> --->
    '<!---#############################################################################################--->
    '<!---#############################################################################################--->
    '<!---#############################################################################################--->
End If
'<!--- end of food consumption  </cfif estimate width --->


   
    
    'end qb
    '############################################################################
    
    
    
    
    
    
    
    
    
    
    
        
    '############################################################################
    '####### end main ###########################################################
    '############################################################################
                
        
    TBL.Edit
    TBL.Fields("lmax").Value = Variable_Length
    TBL.Fields("lmax_type").Value = variable_type1
    
    TBL.Fields("linf").Value = variable_infinity
    TBL.Fields("linf_type").Value = variable_type2
    
    TBL.Fields("K").Value = finalk
    TBL.Fields("PhiPrime").Value = variable_q
    TBL.Fields("to").Value = xto
    
    TBL.Fields("mean_temp").Value = var_temp
    
    TBL.Fields("M").Value = final_mortality
    TBL.Fields("M_se_1st").Value = m1st
    TBL.Fields("M_se_2nd").Value = m2nd
    
    TBL.Fields("life_span").Value = lspan
    
    TBL.Fields("generation_time").Value = GENTIME
    
    TBL.Fields("tm").Value = gtime
    
    TBL.Fields("Lm").Value = vlmaturity
    TBL.Fields("Lm_se_1st").Value = lm_1
    TBL.Fields("Lm_se_2nd").Value = lm_2
    TBL.Fields("Lm_type").Value = maturity_lt
    
    TBL.Fields("Lopt").Value = lmaxyield
    TBL.Fields("Lopt_se_1st").Value = lmaxyield_range1
    TBL.Fields("Lopt_se_2nd").Value = lmaxyield_range2
    TBL.Fields("Lopt_type").Value = yield_lt
    TBL.Fields("Lopt_text").Value = lmaxyield_est
        
        
    TBL.Fields("a").Value = var_a
    TBL.Fields("b").Value = var_b
    TBL.Fields("W").Value = finalval
    TBL.Fields("W_type").Value = finalval_type
    TBL.Fields("LW_length").Value = variable_length2
    TBL.Fields("LW_length_type").Value = variable_type3
    
    
    TBL.Fields("nitrogen").Value = nitrogen
    TBL.Fields("protein").Value = protein
    TBL.Fields("NitrogenProtein_weight").Value = np_weight
    
    TBL.Fields("reproductive_guild").Value = rg
    
    
    TBL.Fields("fecundity").Value = fecundity_v
    TBL.Fields("fecundity_1st").Value = fecundity_v1
    TBL.Fields("fecundity_2nd").Value = fecundity_v2
    'TBL.Fields("fecundity_text").Value = fecundity_text
    
    
    
    TBL.Fields("Emsy").Value = vemsy
    TBL.Fields("Eopt").Value = veopt
    TBL.Fields("Fmsy").Value = vfmsy
    TBL.Fields("Fopt").Value = vfopt
    TBL.Fields("Lc").Value = vLc
    TBL.Fields("Lc_type").Value = Lc_lt
    TBL.Fields("E").Value = vE
    TBL.Fields("YR").Value = vYR
    
    TBL.Fields("resilience").Value = resiliency
    
    
    
    TBL.Fields("rm").Value = Round(vrm, 2)
    TBL.Fields("Lr").Value = vlr
    TBL.Fields("Lr_type").Value = lr_lt
    
    
    TBL.Fields("main_food").Value = mf
    TBL.Fields("trophic_level").Value = tl
         
     
    
    
    TBL.Fields("QB").Value = finalqb
    TBL.Fields("QB_text").Value = finalqb_text
    TBL.Fields("QB_winf").Value = vwinf
    TBL.Fields("QB_temp").Value = var_temp_qb
    TBL.Fields("QB_A").Value = vAfin
       
    
    
    TBL.Update
        
        
    End If
    TBL.MoveNext
Wend


'MsgBox TBL.RecordCount
'MsgBox i
TBL.Close
MDB.Close


End Sub

Private Sub Command4_Click()
'3

Dim MDB As Database
Dim TBL As Recordset

Set MDB = OpenDatabase(App.Path & "\keyfacts.mdb")
Set TBL = MDB.OpenRecordset("SELECT Matrix.*, SPECIES.FamCode, SPECIES.Length, SPECIES.LengthFemale, SPECIES.LTypeMaxM, SPECIES.LTypeMaxF " & _
"FROM Matrix LEFT JOIN SPECIES ON Matrix.SpecCode = SPECIES.SpecCode;", dbOpenDynaset)

Dim i As Long
i = 0
Dim X As Long
Dim SSS As Variant
Dim STR As String

TBL.MoveFirst
While Not TBL.EOF

'start initialize c4
    Variable_Length = Null
    variable_type1 = Null
    variable_infinity = Null
    variable_type2 = Null
    linf_r1 = Null
    Linf_r2 = Null
    finalk = Null
    variable_q = Null
    xto = Null
    var_temp = Null
    final_mortality = Null
    m1st = Null
    m2nd = Null
    lspan = Null
    lspan_r1 = Null
    lspan_r2 = Null
    GENTIME = Null
    gentime_r1 = Null
    gentime_r2 = Null
    gtime = Null
    gtime_r1 = Null
    gtime_r2 = Null
    vlmaturity = Null
    lm_1 = Null
    lm_2 = Null
    maturity_lt = Null
    lmaxyield = Null
    lmaxyield_range1 = Null
    lmaxyield_range2 = Null
    yield_lt = Null
    lmaxyield_est = Null
    var_a = Null
    var_b = Null
    finalval = Null
    finalval_type = Null
    variable_length2 = Null
    variable_type3 = Null
    nitrogen = Null
    protein = Null
    np_weight = Null
    rg = Null
    fecundity_v = Null
    fecundity_v1 = Null
    fecundity_v2 = Null
    fecundity_text = Null
    vemsy = Null
    veopt = Null
    vfmsy = Null
    vfopt = Null
    vLc = Null
    Lc_lt = Null
    vE = Null
    vYR = Null
    resiliency = Null
    vrm = Null
    vlr = Null
    lr_lt = Null
    mf = Null
    tl = Null
    finalqb = Null
    finalqb_text = Null
    vwinf = Null
    var_temp_qb = Null
    vAfin = Null
'end initialize



    If TBL!kf_type = 3 Then
    'If TBL!Speccode = 2 Or TBL!Speccode = 2 Then
        i = i + 1
        X = TBL!SpecCode
        var_tm = TBL!tm_for_KF
        
    STR = "SELECT STOCKS.Stockcode From stocks " & _
    "where stocks.SpecCode=" & X & "and stocks.level = 'species in general'"
    Set getstockcode = MDB.OpenRecordset(STR, dbOpenDynaset)
    If getstockcode.RecordCount <> 0 Then
        vstockcode = getstockcode!StockCode
    Else
        vstockcode = 0
    End If
        
        
        
        
        
    '############################################################################
    '####### start main #########################################################
    '############################################################################
    var_temp = 0
        
    
    linf_r1 = Null
    Linf_r2 = Null
    
    '############################################################################
    'start lmax for KF_type = 1
    Variable_Length = 1
    If TBL!length <> "" And Not IsNull(TBL!length) Then
        Variable_Length = TBL!length
    Else
        If TBL!lengthfemale <> "" And Not IsNull(lengthfemale) Then
            Variable_Length = TBL!lengthfemale
        Else
            Variable_Length = 1
        End If
    End If
    If TBL!ltypemaxm <> "" And Not IsNull(TBL!ltypemaxm) Then
        variable_type1 = TBL!ltypemaxm
    Else
        variable_type1 = ""
    End If
    'end lmax for KF_type = 1
    
        

      
    
    
    
    '############################################################################
    
    
    '############################################################################
    'start linf for KF_type = 3
    
    
        If TBL!length <> "" And Not IsNull(TBL!length) Then
            variable_lmax = TBL!length
            variable_infinity = 10 ^ (0.044 + 0.9841 * (Log(TBL!length) / Log(10)))
            
            
            
    'If X = 502 Then
    'MsgBox TBL!length
    'MsgBox "aa " & variable_infinity
    'End If
    
        Else
            If TBL!lengthfemale <> "" And Not IsNull(TBL!lengthfemale) Then
                variable_lmax = TBL!lengthfemale
                variable_infinity = 10 ^ (0.044 + 0.9841 * (Log(TBL!lengthfemale) / Log(10)))
    'If X = 502 Then
    'MsgBox "bb " & variable_infinity
    'End If
            
            Else
                variable_lmax = 1
                variable_infinity = 10 ^ (0.044 + 0.9841 * (Log(1) / Log(10)))
    
    'If X = 502 Then
    'MsgBox "cc " & variable_infinity
    'End If
            
            End If
        End If
   
    variable_infinity = Round(variable_infinity, 1)
    'If X = 502 Then
    'MsgBox variable_infinity
    'End If
    
    vlmaturity = Round(10 ^ (0.898 * (Log(variable_infinity) / Log(10)) - 0.0781), 1)

    mat_r1 = Round(10 ^ (0.898 * (Log(variable_infinity) / Log(10)) - 0.0781 - 0.127), 1)
    mat_r2 = Round(10 ^ (0.898 * (Log(variable_infinity) / Log(10)) - 0.0781 + 0.127), 1)

      

    'If GetQprime.RecordCount = 0 Then

    If Not IsNull(TBL!length) Then
    
    linf_r1 = Round(10 ^ (0.044 + 0.9841 * (Log(TBL!length) / Log(10)) - 0.074), 1)
    Linf_r2 = Round(10 ^ (0.044 + 0.9841 * (Log(TBL!length) / Log(10)) + 0.074), 1)
    
    End If
    
    'End If
    
    
    
    'end linf for KF_type = 1
    '############################################################################
    
    
    '############################################################################
    'start K for KF_type = 3
    
    'this one is for kf_type = 1
    'finalk = (10 ^ (variable_q - 2 * (Log(variable_infinity) / Log(10))))
    
    
xxxk = -Log(1 - vlmaturity / variable_infinity) / (var_tm)
xto = -1 * (10 ^ (-0.3922 - 0.2752 * (Log(variable_infinity) / Log(10)) - 1.038 * (Log(xxxk) / Log(10))))

'If X = 502 Then
'    MsgBox xxxk
'    MsgBox xto
'End If

xto = Round(xto, 2)
finalk = -Log(1 - vlmaturity / variable_infinity) / (var_tm - xto)
finalk = Round(finalk, 2)
    
    
    
    
    
    
    

    'end K for KF_type = 3
    '############################################################################

    
    
    
    
    '############################################################################
    'start var_temp
    
    If var_temp = 0 Then
        STR = "SELECT Avg(POPGROWTH.Temperature) AS AvgOfT " & _
        "From POPGROWTH " & _
        "WHERE (((POPGROWTH.SpecCode)=" & TBL!SpecCode & ") AND (Not (POPGROWTH.Temperature)=0 " & _
        "And (POPGROWTH.Temperature) Is Not Null))"
        Set avgt = MDB.OpenRecordset(STR, dbOpenDynaset)
        
        While Not avgt.EOF
            var_temp = avgt!avgoft
            avgt.MoveNext
        Wend
    End If
    
    
    
    
    If var_temp = 0 Then
    STR = "SELECT Avg(([tempmin]+[tempmax])/2) AS vAvg " & _
    "From STOCKS " & _
    "WHERE   (   ((stocks.tempmin)<>0 And (stocks.tempmin) Is Not Null) AND " & _
                "((stocks.tempmax)<>0 And (stocks.tempmax) Is Not Null) AND " & _
                "((stocks.Speccode)=" & TBL!SpecCode & ")    )"
    Set opt_stocks = MDB.OpenRecordset(STR, dbOpenDynaset)
    If opt_stocks!vAvg > 0 Then
        var_temp = opt_stocks!vAvg
    End If
    End If
    
    
If IsNull(var_temp) Then
    STR = "SELECT STOCKS.EnvTemp From STOCKS " & _
    "WHERE (((STOCKS.stockcode)=" & vstockcode & "));"
    Set opt_stocks = MDB.OpenRecordset(STR, dbOpenDynaset)
    
    While Not opt_stocks.EOF
        If opt_stocks!envtemp = "boreal" Then
            var_temp = 6
        ElseIf opt_stocks!envtemp = "deep-water" Then
            var_temp = 8
        ElseIf opt_stocks!envtemp = "high altitude" Then
            var_temp = 1
        ElseIf opt_stocks!envtemp = "polar" Then
            var_temp = 1
        ElseIf opt_stocks!envtemp = "subtropical" Then
            var_temp = 17
        ElseIf opt_stocks!envtemp = "temperate" Then
            var_temp = 10
        ElseIf opt_stocks!envtemp = "tropical" Then
            var_temp = 25
        End If
        opt_stocks.MoveNext
    Wend
End If
   
    
    'end var_temp
    '############################################################################
    
    
    '############################################################################
    'start variable_type2
    
    
    STR = "SELECT POPgrowth.Speccode,popgrowth.temperature From POPgrowth " & _
    "WHERE (((POPgrowth.Speccode)=" & X & "))"
    Set withgrowth = MDB.OpenRecordset(STR, dbOpenDynaset)
    
    'If withgrowth.RecordCount = 0 Then
    '    If variable_type1 <> "" Then
    '        variable_type2 = variable_type1
    '    Else
    '        variable_type2 = ""
    '    End If
    'Else
    '    variable_type2 = xxxVT2
    'End If
    'If withgrowth.RecordCount = 0 Then
    '    variable_type2 = variable_type1
    'Else
    '    variable_type2 = variable_type2
    'End If
    
    variable_type2 = variable_type1
    
    
    
    'end variable_type2
    '############################################################################
    
    
    '############################################################################
    'start M and others
    final_mortality = 10 ^ (0.333 - 0.246 * (Log(variable_infinity) / Log(10)) + 0.744 * (Log(finalk) / Log(10)) + 0.01 * var_temp)
    m1st = Null
    m2nd = Null
    
    If variable_type2 = "TL" Then
        m1st = 10 ^ (0.333 - 0.246 * (Log(variable_infinity) / Log(10)) + 0.744 * (Log(finalk) / Log(10)) + 0.01 * var_temp - 0.18)
        m2nd = 10 ^ (0.333 - 0.246 * (Log(variable_infinity) / Log(10)) + 0.744 * (Log(finalk) / Log(10)) + 0.01 * var_temp + 0.18)
    Else
        'm1st = ""
        'm2nd = ""
    End If
    'end M and others
    '############################################################################
    
    
    '############################################################################
    'start life span
    lspan = (3 / finalk) + xto
    
    
    If Not IsNull(linf_r1) Then
k1 = Round((-Log(1 - vlmaturity / linf_r1) / (var_tm - xto)), 2)
k2 = Round((-Log(1 - vlmaturity / Linf_r2) / (var_tm - xto)), 2)
lspan_r1 = Round((3 / k1) + xto, 1)
lspan_r2 = Round((3 / k2) + xto, 1)
    End If
    
    
    
    
    
    'end life span
    '############################################################################
    
    
    '############################################################################
    'start gentime
    
    
'STR = "SELECT popgrowthref,loo,k,type,lm,tlinfinity From POPGROWTH " & _
'"WHERE (((POPGROWTH.SpecCode)=" & X & ")) order by popgrowth.loo,popgrowth.type;"
'Set medgrowth = MDB.OpenRecordset(STR, dbOpenDynaset)
'If medgrowth.RecordCount = 0 Then
'Else
'    vlmaturity = 10 ^ (0.898 * (Log(variable_infinity) / Log(10)) - 0.0781)
'    If (xxxLoo) Then '<!--- meaning with growth, median Qprime --->
'        qtest1 = variable_infinity * (3 / (3 + final_mortality / finalk))
'        qtest2 = 10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742)
'        If qtest1 < qtest2 Then
'            varlopt = 10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742)
'        Else
'            varlopt = variable_infinity * (3 / (3 + final_mortality / finalk))
'        End If
'    Else
'        varlopt = 10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742)
'    End If
'    evlmaturity = 10 ^ (0.898 * (Log(variable_infinity) / Log(10)) - 0.0781)
'    If evlmaturity >= varlopt Then
'        lm100 = evlmaturity + (variable_infinity - evlmaturity) / 4
'        gentime = xto + (-1 * (Log(1 - Val(lm100) / Val(variable_infinity)) / Val(finalk)))
'    Else
'        gentime = xto + (-1 * (Log(1 - Val(varlopt) / Val(variable_infinity)) / Val(finalk)))
'    End If
'End If
        
        
        
varlopt = Round(10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742), 1)

evlmaturity = Round(10 ^ (0.898 * (Log(variable_infinity) / Log(10)) - 0.0781), 1)
If evlmaturity >= varlopt Then
    lm100 = evlmaturity + (variable_infinity - evlmaturity) / 4
    GENTIME = xto + (-1 * (Log(1 - lm100 / variable_infinity) / finalk))
    
    'If X = 502 Then
    '    MsgBox "111222"
    'End If
Else
       
lmaxyield_range1 = Round(10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742 - 0.073), 1)
lmaxyield_range2 = Round(10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742 + 0.073), 1)

       
       
    GENTIME = xto + (-1 * (Log(1 - varlopt / variable_infinity) / finalk))
    gentime_r1 = xto + (-1 * (Log(1 - lmaxyield_range1 / variable_infinity) / finalk))
    gentime_r2 = xto + (-1 * (Log(1 - lmaxyield_range2 / variable_infinity) / finalk))
    
    'If X = 502 Then
    '    MsgBox "333444"
    'End If

End If

        
        
        
        
        
        
    
    'end gentime
    '############################################################################
    
    '############################################################################
    'start tm
    
    'for kf_type=1
    'gtime = Round(xto, 2) + (-1 * (Log(1 - Val(Round(vlmaturity, 1)) / Val(Round(variable_infinity, 2))) / Val(Round(finalk, 2))))
    
    
    gtime = Round(xto + (-1 * (Log(1 - vlmaturity / variable_infinity) / finalk)), 1)
    
    'If X = 845 Then
    '    MsgBox xto & " " & vlmaturity & " " & variable_infinity & " " & finalk
    'End If
    
    
    If Not IsNull(linf_r1) Then
    gtime_r1 = Round(xto + (-1 * (Log(1 - mat_r1 / linf_r1) / finalk)), 1)
    gtime_r2 = Round(xto + (-1 * (Log(1 - mat_r2 / Linf_r2) / finalk)), 1)
    End If
    
    


    
    
    
    
    'If TBL!speccode = 2 Then
    'MsgBox Round(xto, 2)
    'MsgBox Round(vlmaturity, 1)
    'MsgBox Round(variable_infinity, 2)
    'MsgBox Round(finalk, 2)
    'End If
    
    'end tm
    '############################################################################
        
        
        
    '############################################################################
    'start Lm se
    lm_1 = 10 ^ (0.898 * (Log(variable_infinity) / Log(10)) - 0.0781 - 0.127)
    lm_2 = 10 ^ (0.898 * (Log(variable_infinity) / Log(10)) - 0.0781 + 0.127)
    maturity_lt = variable_type2
    'end Lm se
    '############################################################################
    
        
    '############################################################################
    'start lopt
    
    If xxxLoo <> 0 Then '<!--- meaning with growth, median Qprime --->
        qtest1 = variable_infinity * (3 / (3 + final_mortality / finalk))
        qtest2 = 10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742)
        If qtest1 < qtest2 Then
            lmaxyield = 10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742)
            'MsgBox "2"
        Else
            lmaxyield = variable_infinity * (3 / (3 + Round(final_mortality, 2) / Round(finalk, 2)))
            'MsgBox "3"
            'MsgBox lmaxyield
            'MsgBox variable_infinity
            'MsgBox final_mortality
            'MsgBox finalk
        End If
    Else
        lmaxyield = 10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742)
        'MsgBox "4"
    End If
    
    
    'this is from kf_type=1
    'If xxxLoo = 0 Then '<!--- meaning w/o growth, median Qprime --->
    '    lmaxyield_est = "Estimated from Linf."
    'Else
    
        'ito ay parang pang kf_type=1
        'If qtest1 < qtest2 Then
        '    lmaxyield_range1 = 10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742 - 0.073)
        '    lmaxyield_range2 = 10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742 + 0.073)
        '    lmaxyield_est = "Estimated from Linf."
        'Else
        '    lmaxyield_range1 = Null
        '    lmaxyield_range2 = Null
        '    lmaxyield_est = "Estimated from Linf., K and M."
        'End If
    'End If
    
    
    
    
    
    
        
    yield_lt = variable_type2
    
    'end lopt
    '############################################################################
    
    
    '############################################################################
    'start of l-w
    variable_length2 = Round(variable_infinity, 1)


    'start <!--- get the median of LW --->
    STR = "SELECT  POPLW.SpecCode,POPLW.LengthMin,POPLW.LengthMax,POPLW.Number,POPLW.Sex, POPLW.a, POPLW.b, COUNTREF.paese, " & _
        "poplw.autoctr, poplw.locality,poplw.type, poplw.a , poplw.b FROM COUNTREF INNER JOIN POPLW ON COUNTREF.C_Code = POPLW.C_Code " & _
    "WHERE (((POPLW.SpecCode)=" & TBL!SpecCode & "))        order by poplw.b"
    Set medlwb = MDB.OpenRecordset(STR, dbOpenDynaset)
    
    If medlwb.RecordCount <> 0 Then
        xLW = (medlwb.RecordCount / 2) + 0.5
        xLW = Round(xLW)
        
        'If X = 68 Then
        '    MsgBox xLW
        'End If
        
        ii = 0
        While Not medlwb.EOF
            ii = ii + 1
            If ii = xLW Then
                v_a = medlwb!a
                v_b = medlwb!b
                lwtype = medlwb!Type
            End If
            medlwb.MoveNext
        Wend
    End If
    'end <!--- get the median of LW --->



'<!--- start of lw --->
If medlwb.RecordCount <> 0 Then
    'start parang hindi dada-anan ito
    'if #parameterexists(variable_length2)# is "No">
    '    variable_length2 = #numberformat(length,"9999.9")#>
    'end if
    'end parang hindi dada-anan ito
    
    
    'if #parameterexists(var_a)# is "No">
    var_a = v_a
    'end if
    'if #parameterexists(var_b)# is "No">
    var_b = v_b
    'end if
    finalval = ((variable_length2 ^ Round(var_b, 3)) * Round(var_a, 4))
    'finalval = ((46 ^ 3.13) * 0.0054)
    
    'If X = 68 Then
    '    MsgBox variable_length2 & " " & var_b & " " & var_a & " = " & finalval
    'End If
    
    variable_type3 = lwtype
    eli01 = Len(Trim(Round(finalval, 1))) * 2


    'If TBL!speccode = 2 Then
    '    MsgBox var_a
    '    MsgBox var_b
    '    MsgBox finalval
    'End If

    
    np_weight = finalval
    
    If finalval > 200 Then
        If finalval > 20000 Then
            finalval = Round(finalval / 1000, 1)
            finalval_type = "kg"
        Else
            finalval = Round(finalval)
            finalval_type = "g"
        End If
    Else
        finalval = Round(finalval, 1)
        finalval_type = "g"
    End If
    
       
    'W = finalval
    'a = Var_a
    'b = Var_b
   
    'Nitrogen & protein:
    
    '<!--- start new weight --->
    If finalval <> 0 Then
        'NP_weight = finalval   .... transferred up to preserve the decimal places..
        
        'If Trim(finalval_type) = "kg" Then
        '    NP_weight = NP_weight * 1000
        'End If
        
        
        
        nitrogen = 10 ^ (1.03 * (Log(np_weight) / Log(10)) - 1.65)
        nitrogen = Round(nitrogen, 1)
        protein = Round(6.25 * nitrogen, 1)
        np_weight = Round(np_weight)
        
        
        
    Else
        np_weight = Null
        nitrogen = Null
        protein = Null
    End If
    '<!--- end new weight --->
    
Else
    '<!--- don't show --->
End If

'<!--- end of lw --->
    
    
    'end of l-w
    '############################################################################
    
    
    '############################################################################
    'start reproductive_guild
    rg = ""
    If withgrowth.RecordCount <> 0 Then
    
    'MsgBox withgrowth.RecordCount
    'MsgBox "riz = " & vstockcode
    
    If withgrowth.RecordCount <> 0 Then
    
    
    
    STR = "SELECT REPRODUC.RepGuild1, REPRODUC.RepGuild2 From REPRODUC " & _
    "WHERE (((REPRODUC.StockCode)=" & vstockcode & "))"
    Set getguild = MDB.OpenRecordset(STR, dbOpenDynaset)
    rg = ""
    If getguild.RecordCount <> 0 Then
        If Trim(getguild!repguild1) = "" Then
            rg = " "
        Else
            rg = getguild!repguild1
        End If
        rg = rg + ": "
        If Trim(getguild.repguild2) = "" Then
            rg = rg & " "
        Else
            rg = rg & getguild!repguild2
        End If
        rg = Trim(rg)
    End If
    End If
    End If
    'end reproductive_guild
    '############################################################################
    

    
    
    
    '############################################################################
    'start fecundity
    
    fecundity_v = Null
    fecundity_v1 = Null
    fecundity_v2 = Null
    fecundity_text = Null
    
    
    
    If withgrowth.RecordCount <> 0 And vstockcode <> 0 Then
    
    'If X = 69 Then
    'MsgBox X
    'MsgBox Vstockcode
    'End If

STR = "SELECT Min(SPAWNING.FecundityMin) AS MinOfFecundityMin From SPAWNING " & _
"GROUP BY SPAWNING.StockCode, SPAWNING.SpecCode " & _
"HAVING ((Not (Min(SPAWNING.FecundityMin))=0) AND ((SPAWNING.StockCode)=" & vstockcode & ") AND ((SPAWNING.SpecCode)=" & X & ")) " & _
"ORDER BY Min(SPAWNING.FecundityMin)"
Set getmin = MDB.OpenRecordset(STR, dbOpenDynaset)

STR = "SELECT Max(SPAWNING.FecundityMax) AS MaxOfFecundityMax From SPAWNING GROUP BY SPAWNING.StockCode, SPAWNING.SpecCode " & _
"HAVING (((SPAWNING.StockCode)=" & vstockcode & ") AND ((SPAWNING.SpecCode)=" & X & ") AND (Not (Max(SPAWNING.FecundityMax))=0))"
Set getmax = MDB.OpenRecordset(STR, dbOpenDynaset)
    
    
If getmin.RecordCount <> 0 Or getmax.RecordCount <> 0 Then
    
    y01 = 0
    y02 = 0
    fecundity_v1 = ""
    fecundity_v2 = ""
    
    
    If getmin.RecordCount <> 0 And getmax.RecordCount <> 0 Then
    If getmin!minoffecunditymin > 0 And getmax!maxoffecunditymax > 0 Then
    y01 = Log(getmin.minoffecunditymin) / Log(10)
    y02 = Log(getmax.maxoffecunditymax) / Log(10)
    y01 = y01 + y02
    y01 = y01 / 2
    fecundity_v = Round(10 ^ y01)
    End If
    End If
    
    If getmin.RecordCount <> 0 Then
        If getmin!minoffecunditymin = "" Then
            fecundity_v1 = "no value (min.)"
        Else
            fecundity_v1 = Round(getmin!minoffecunditymin)
        End If
    Else
        fecundity_v1 = "no record (min.)"
    End If
    
    
    If getmax.RecordCount <> 0 Then
        If getmax!maxoffecunditymax = "" Then
            fecundity_v2 = "no value (max.)"
        Else
            fecundity_v2 = Round(getmax!maxoffecunditymax)
        End If
    Else
        fecundity_v2 = "no record (max.)"
    End If
    
    
    
    If getmin.RecordCount <> 0 And getmax.RecordCount <> 0 Then
        fecundity_text = "Estimated as geometric mean."
    End If


End If
    End If
    'end fecundity
    '############################################################################
    
    
    
'############################################################################
'start yrecruit

'<!--- start yrecruit --->

    vemsy = Null
    veopt = Null
    vfmsy = Null
    vfmsy_rm = Null
    vfopt = Null
    vLc = Null
    Lc_lt = Null
    vE = Null
    vYR = Null





vE = 0.5
vLc = Round(0.4 * variable_infinity, 1)
If final_mortality > 0 Then
    
    vU = 1 - vLc / Round(variable_infinity, 2)
    MK = Round(final_mortality, 2) / Round(finalk, 2)
    
    vx1 = 0
    vx2 = 0
    firstloop = "Y"
    
    oldy = 0
    vlope = 0
    vemsy = 0
    veopt = 0
    vfmsy = 0
    vfmsy_rm = 0
    vfopt = 0
    
    
    While vx1 <= 1
        
        vx1 = vx2
        vx2 = vx2 + 0.001

        vm1 = (1 - vx1) / MK
        vy1 = vx1 * (vU ^ MK) * (1 - ((3 * vU) / (1 + vm1)) + ((3 * vU ^ 2) / (1 + 2 * vm1)) - ((vU ^ 3) / (1 + 3 * vm1)))
        vm2 = (1 - vx2) / MK
        vy2 = vx2 * (vU ^ MK) * (1 - ((3 * vU) / (1 + vm2)) + ((3 * vU ^ 2) / (1 + 2 * vm2)) - ((vU ^ 3) / (1 + 3 * vm2)))
        vslope = (vy2 - vy1) / (vx2 - vx1)
        
        If oldy <> 0 Then
            If vy1 >= oldy Then
            Else
                vemsy = Round(vx1 - 0.001, 2)
                '<cfbreak>
                GoTo EliGo
            End If
        End If
        
        
        oldy = vy1
        
        If firstloop = "Y" Then
            firstvalue = (vy2 - vy1) / (vx2 - vx1)
            firstloop = "N"
        End If

        
    
        If veopt = 0 Then
        If Round(vslope, 3) = Round(firstvalue / 10, 3) Then
            veopt = vx1
        End If
        End If
        
        
    Wend
EliGo:
    '<!--- end get e  --->
    

    vm = (1 - vE) / MK
    lijosh = vE * (vU ^ MK) * (1 - ((3 * vU) / (1 + vm)) + ((3 * vU ^ 2) / (1 + 2 * vm)) - ((vU ^ 3) / (1 + 3 * vm)))
    '<input type="text" name="vYR" value=round(lijosh,'9.9999')#" size="6" onFocus="noedit(2)"  align="right">
    vYR = Round(lijosh, 4)

Else
    '<input type="text" name="vYR" value="" size="6" onFocus="noedit(2)"  align="right">
    vYR = Null
End If

    vLc = Round(vLc, 1)
    Lc_lt = Trim(variable_type2)
    e = vE = Round(vE, 2)
    
    If final_mortality > 0 Then
    
        veopt = Round(veopt, 2)
        vemsy = Round(vemsy, 2)
        'vfmsy = Round(final_mortality * vEmsy / (1 - vEmsy), 2)
        vfmsy = final_mortality * vemsy / (1 - vemsy)
        vfmsy_rm = final_mortality * vemsy / (1 - vemsy)
        'vFopt = Round(final_mortality * vEopt / (1 - vEopt), 2)
        vfopt = final_mortality * veopt / (1 - veopt)
        
        
        
        
        
        
        veopt = Round(veopt, 2)
        vemsy = Round(vemsy, 2)
        vfmsy = Round(vfmsy, 2)
        vfopt = Round(vfopt, 2)
    Else
        vemsy = Null
        veopt = Null
        vfmsy = Null
        vfmsy_rm = Null
        vfopt = Null
    End If





'<!--- end yrecruit   width  msy --->


'end yrecruit
'############################################################################
    
    
    '############################################################################
    'start resiliency
    resiliency = Null
    
       
    If Round(finalk, 2) <= 0.05 Then
        resiliency = "Very low; decline threshold 0.70"
    ElseIf Round(finalk, 2) <= 0.15 Then
        resiliency = "Low; decline threshold 0.85"
    ElseIf Round(finalk, 2) <= 0.3 Then
        resiliency = "Medium; decline threshold 0.95"
    ElseIf Round(finalk, 2) > 0.3 Then
        resiliency = "High; decline threshold 0.99"
    Else
        resiliency = "Please enter values for K."
    End If
    
      
    
    
    'end resiliency
    '############################################################################
    
    
    
    
    
    
    '############################################################################
    'start rm
    vlr = Round(0.4 * variable_infinity, 1)
    vfmsy = Round(vfmsy, 2)
    vrm = Round(2 * vfmsy, 2)
    vypdt = Log(2) / vrm
    lr_lt = Trim(variable_type2)
    
    'If X = 2 Then
    '    MsgBox vfmsy
    '    MsgBox vrm
    'End If
    
    'end rm
    '############################################################################
    
    
    
    '############################################################################
    'start eco
    STR = "SELECT  ECOLOGY.StockCode,ECOLOGY.DietTroph,ECOLOGY.DietSeTroph,ECOLOGY.FoodTroph, " & _
    "ECOLOGY.FoodSeTroph,0 as EcoTroph,0 as EcoSeTroph,'' as mainfood,ECOLOGY.herbivory2 " & _
    "From ECOLOGY WHERE (((ECOLOGY.StockCode)=" & vstockcode & "))"
    Set witheco = MDB.OpenRecordset(STR, dbOpenDynaset)
    
mf = Null
tl = Null
If witheco.RecordCount <> 0 Then
    If witheco!mainfood <> "" Then
        mf = witheco!mainfood
    End If
    If witheco!herbivory2 <> "" Then
        mf = mf & " " & witheco!herbivory2
    End If
    
    
    'Trophic level:
    tl = Null
    If witheco!dietTroph <> "" Then
        tl = Round(witheco!dietTroph, 1)
        If witheco!DietSeTroph <> 0 Then
            tl = tl & "&nbsp;&nbsp;&nbsp;"
            tl = tl & "+/- s.e. " & Round(witheco!DietSeTroph, 2)
        End If
        tl = tl & " Estimated from diet data."
    Else
        If witheco!foodTroph <> "" Then
            tl = tl & " " & Round(witheco!foodTroph, 1)
            If witheco!FoodSeTroph <> 0 Then
                tl = tl & "+/- s.e. " & Round(witheco!FoodSeTroph, 2)
            End If
            tl = tl & " Estimated from food data."
        Else
            If witheco!EcoTroph <> "" Then
                tl = tl & " " & Round(witheco!EcoTroph, 1)
                If witheco!EcoSeTroph <> 0 Then
                    tl = tl & "+/- s.e. " & Round(witheco!EcoSeTroph, 2)
                End If
                tl = tl & " Estimated from Ecopath model."
            End If
        End If
    End If
End If


'<!---start of food consumption  </cfif --->

    
    'end eco
    '############################################################################
    
    
    
    
    
    
    
    
    '############################################################################
    'start qb
    
    finalqb = Null
    finalqb_text = Null
    vwinf = Null
    var_temp_qb = Null
    vAfin = Null
    
    
    
    
    
'<!--- start food consumption --->
contqb = "Y"
'<!--- start get median popqb --->
STR = "SELECT popqb.popqb FROM popqb WHERE (((popqb.speccode)=" & X & ")) ORDER BY popqb.popqb"
Set qbmedian = MDB.OpenRecordset(STR, dbOpenDynaset)


'start of old
'vmedian = (qbmedian.RecordCount / 2) + 0.5
'vmedian = Round(vmedian)
'end of old


vMedian = (qbmedian.RecordCount / 2) + 0.5
int_part = Int(vMedian)
str_conv = "" & vMedian

If Len(str_conv) = 1 Then
    vMedian = Round(vMedian)
Else

    If Mid(str_conv, Len(str_conv) - 1, 2) = ".5" Then
        vMedian = int_part + 1
    Else
        vMedian = Round(vMedian)
    End If
End If




vpopqb = 0


'If X = 69 Then
'    MsgBox qbmedian.RecordCount
'    MsgBox (qbmedian.RecordCount / 2) + 0.5
'    MsgBox Int((qbmedian.RecordCount / 2) + 0.5)
'    MsgBox Round(vmedian)
'    MsgBox Round(vmedian, 1)
'End If


ii = 0
While Not qbmedian.EOF

    ii = ii + 1
    If ii = vMedian Then
        vpopqb = qbmedian!popqb
    End If
    qbmedian.MoveNext
Wend


'<!--- end get median popqb --->

If vpopqb > 0 Then
    explain = "with popqb record"
    contqb = "N"
Else
    explain = "no popqb record"
    
    '<!---#############################################################################################--->
    '<!---### start of no popqb #######################################################################--->
    '<!---#############################################################################################--->
    If medlwb.RecordCount# > 0 Then
        explain = explain & "; with lw rel"
    Else
        explain = explain & "; no lw rel"
    End If

    '<!--- start of A --->
    STR = "SELECT Swimming.AspectRatio as aspectratio From Swimming WHERE (((Swimming.SpecCode)=" & X & "));"
    Set getAR = MDB.OpenRecordset(STR, dbOpenDynaset)
    vAfin = 0
    If getAR.RecordCount > 0 Then
       While Not getAR.EOF
            If getAR!aspectratio > 0 Then
                vAfin = getAR!aspectratio
                explain = explain & "; with aspect ratio"
            Else
                vAfin = 0
                explain = explain & "; w/o aspect ratio"
            End If
            getAR.MoveNext
        Wend
    Else
        explain = explain & "; w/o aspect ratio"
    End If
    '<!--- end of A --->
    
    
    '<!--- start of h and d --->
    SelectHD = "N"
    '<!---
    '    go ecology feeding type
    '    if none troph diettroph
    '        if none troph foodtroph items
    '            if none troph ecotroph
    '--->


    STR = "SELECT ECOLOGY.Herbivory2, 0 as EcoTroph, ECOLOGY.DietTroph, ECOLOGY.FoodTroph From ECOLOGY " & _
    "WHERE (((ECOLOGY.StockCode)=" & vstockcode & "))"
    Set getft = MDB.OpenRecordset(STR, dbOpenDynaset)

    
    
    If getft.RecordCount <> 0 Then
    If Trim(getft!herbivory2) <> "" Then
    
        If Trim(getft!herbivory2) = "mainly animals (troph. 2.8 and up)" Then
            vh = 0
            vd = 0
        ElseIf Trim(getft!herbivory2) = "mainly animals (troph. 2.8 up)" Then
            vh = 0
            vd = 0
        ElseIf Trim(getft!herbivory2) = "mainly plants/detritus (troph. 2-2.19)" Then
            vh = 0
            vd = 1
        ElseIf Trim(getft.herbivory2) = "plants/detritus+animals (troph. 2.2-2.79)" Then
            vh = 1
            vd = 0
        End If
        explain = explain & "; w/ feeding type"
        WithFT = 1
    
    Else
    
    
        explain = explain & "; w/o feeding type"
        WithFT = 0
        If (getft!dietTroph) > 0 Then
            If getft!dietTroph >= 2 And getft!dietTroph <= 2.19 Then
                vh = 0
                vd = 1
            ElseIf getft!dietTroph >= 2.2 And getft!dietTroph <= 2.79 Then
                vh = 1
                vd = 0
            ElseIf getft!dietTroph >= 2.8 Then
                vh = 0
                vd = 0
            End If
            explain = explain & "; from diettroph"
        Else
            If getft!foodTroph > 0 Then
                If getft!foodTroph >= 2 And getft.foodTroph <= 2.19 Then
                    vh = 0
                    vd = 1
                ElseIf getft.foodTroph >= 2.2 And getft.foodTroph <= 2.79 Then
                    vh = 1
                    vd = 0
                ElseIf getft!foodTroph >= 2.8 Then
                    vh = 0
                    vd = 0
                End If
                explain = explain & "; from foodtroph"
            Else
                If getft!EcoTroph > 0 Then
                    If getft!EcoTroph >= 2 And getft!EcoTroph <= 2.19 Then
                        vh = 0
                        vd = 1
                    ElseIf getft!EcoTroph >= 2.2 And getft!EcoTroph <= 2.79 Then
                        vh = 1
                        vd = 0
                    ElseIf getft!EcoTroph >= 2.8 Then
                        vh = 0
                        vd = 0
                    End If
                    explain = explain & "; from ecotroph"
                Else
                    explain = explain & "; no diet,food,eco trophs; select h,d"
                    '<!--- blank means yes
                    'contqb = "Y">
                    '--->
                    SelectHD = "Y"
                    '<!---
                    'contqb = "N">
                    'cont. with the search
                    'now the genus of median of diet,food,eco
                    '--->
                End If
            End If
        End If
    End If
    End If
    '<!--- end of h and d --->

'<!---#############################################################################################--->
'<!---### end of no popqb #########################################################################--->
'<!---#############################################################################################--->
End If
'<!--- end food consumption --->









'<!---start of food consumption  </cfif --->

'<td align="left">Food consumption (Q/B):</td>
If vpopqb > 0 Then
    finalqb = Round(vpopqb, 2)
    finalqb_text = "times the body weight per year"
Else
    '<!---#############################################################################################--->
    '<!---#############################################################################################--->
    '<!---#############################################################################################--->
    If contqb = "Y" Then
        '<!--- start for winf --->
        If medlwb.RecordCount > 0 Then
            'orig
            'vwinf = finalval
            
            'use this instead as the finalval is sometimes in kg
            vwinf = np_weight
            'If X = 68 Then
            '    MsgBox "111"
            'End If
            
            whereWinf = "lw"
        Else
            vwinf = 0.01 * variable_infinity ^ 3
            'If X = 68 Then
            '    MsgBox "222"
            'End If
            
            whereWinf = "DP"
        End If
        '<!--- end for winf --->
        If SelectHD = "Y" Then
            vh = 3
            vd = 3
        End If
        'start vb comment
        '<input type="hidden" name="vh" value="#vh#" size="1" onFocus="noedit(2)"  >
        '<input type="hidden" name="vd" value="#vd#" size="1" onFocus="noedit(2)"  >
        'end vb comment
        
        If vAfin > 0 Then
            xyz = vAfin
        Else
            xyz = 1.32
        End If
        If vh = 3 Then
            '<!---w/o h, d --->
            elix = 10 ^ (7.964 - 0.204 * Log(vwinf) / Log(10) - 1.965 * (1000 / (var_temp + 273.15)) + 0.083 * xyz + 0.532 * 1 + 0.398 * 0)
            elix2 = 10 ^ (7.964 - 0.204 * Log(vwinf) / Log(10) - 1.965 * (1000 / (var_temp + 273.15)) + 0.083 * xyz + 0.532 * 0 + 0.398 * 0)
            elix = (elix2 + elix) / 2
        Else
            '<!---with h, d--->
            
            'old
            'elix = 10 ^ (7.964 - 0.204 * Log(vwinf) / Log(10) - 1.965 * (1000 / (var_temp + 273.15)) + 0.083 * xyz + 0.532 * vh + 0.398 * vd)
            
            
            'start new
            If vwinf <> Null Then
                elix = 10 ^ (7.964 - 0.204 * Log(vwinf) / Log(10) - 1.965 * (1000 / (var_temp + 273.15)) + 0.083 * xyz + 0.532 * vh + 0.398 * vd)
            Else
                elix = 0
            End If
            'end new
            
        End If
        
        
        If whereWinf = "none" Then
            finalqb = Null
        Else
            finalqb = Round(elix, 1)
        End If
        finalqb_text = "times the body weight per year"
        
        
        'Enter Winf, temperature, aspect ratio (A), and food type to estimate Q/B
        If whereWinf = "none" Then
            'Winf =
            vwinf = Null
            'If X = 68 Then
            '    MsgBox "333"
            'End If
            
            'Temp. =
            var_temp_qb = Round(var_temp, 1)
        Else
            'Winf =
            vwinf = Round(vwinf, 1)
            
            'If X = 68 Then
            '    MsgBox "444"
            'End If
            
            'Temp. =
            var_temp_qb = Round(var_temp, 1)
        End If
        
        If vAfin > 0 Then
            'A =
            vAfin = Round(vAfin, 2)
        Else
            vAfin = 1.32
            'A =
            vAfin = Round(vAfin, 2)
            
            '<!---eeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeee--->
            
            'start vb comment
            '<td colspan="8"><img src="../jpgs/Tails.gif" height=29 border=0 alt=""></td>
            '<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
            '<td><input type="radio" name="eli"          onClick="finc('#contqb#',6.55)"></td>
            '<td><input type="radio" name="eli"          onClick="finc('#contqb#',1.9)"></td>
            '<td><input type="radio"  checked name="eli" onClick="finc('#contqb#',1.32)"></td>
            '<td><input type="radio" name="eli"          onClick="finc('#contqb#',1.32)"></td>
            '<td><input type="radio" name="eli"          onClick="finc('#contqb#',1.63)"></td>
            'end vb comment
            
        End If
        If SelectHD = "Y" Then
            'start vb comment
            '<input type="hidden" name="omni" value="1">
            '<td align="center"><font size="2">Detrivore</font></td>
            '<td align="center"><font size="2">Herbivore</font></td>
            '<td align="center"><font size="2">Omnivore</font></td>
            '<td align="center"><font size="2">Carnivore</font></td>
            '<td align="center"><input type="radio" name="riz"           onClick="changeHD('#contqb#',0,1)"></td>
            '<td align="center"><input type="radio" name="riz"          onClick="changeHD('#contqb#',1,0)"></td>
            '<td align="center"><input type="radio" name="riz" checked  onClick="changeHD('#contqb#',3,3)"></td>
            '<td align="center"><input type="radio" name="riz"           onClick="changeHD('#contqb#',0,0)"></td>
            '</tr></table>
            'end vb comment
            
        Else
            'start vb comment
            '<input type="hidden" name="omni" value="0">
            '<td align="center"><font size="2">Detrivore</font></td>
            '<td align="center"><font size="2">Herbivore</font></td>
            '<td align="center"><font size="2">Omnivore</font></td>
            '<td align="center"><font size="2">Carnivore</font></td>
            '<td align="center"><input type="radio" name="riz" if #vh# is 0 and #vd# is 1>checkedend if onClick="changeHD('#contqb#',0,1)"></td>
            '<td align="center"><input type="radio" name="riz" if #vh# is 1 and #vd# is 0>checkedend if onClick="changeHD('#contqb#',1,0)"></td>
            '<td align="center"><input type="radio" name="riz"                                              onClick="changeHD('#contqb#',3,3)"></td>
            '<td align="center"><input type="radio" name="riz" if #vh# is 0 and #vd# is 0>checkedend if onClick="changeHD('#contqb#',0,0)"></td>
            'end vb comment
        End If
    End If '<!--- if #contqb# is "Y"> --->
    '<!---#############################################################################################--->
    '<!---#############################################################################################--->
    '<!---#############################################################################################--->
End If
'<!--- end of food consumption  </cfif estimate width --->


   
    
    'end qb
    '############################################################################
    
    
    
    
    
    
    
    
    
    
    
        
    '############################################################################
    '####### end main ###########################################################
    '############################################################################
                
        
    TBL.Edit
    TBL.Fields("lmax").Value = Variable_Length
    TBL.Fields("lmax_type").Value = variable_type1
    
    TBL.Fields("linf").Value = variable_infinity
    TBL.Fields("linf_type").Value = variable_type2
    
    TBL.Fields("linf_1st").Value = linf_r1
    TBL.Fields("linf_2nd").Value = Linf_r2
    
   
    
    TBL.Fields("K").Value = finalk
    TBL.Fields("PhiPrime").Value = variable_q
    TBL.Fields("to").Value = xto
    
    TBL.Fields("mean_temp").Value = var_temp
    
    TBL.Fields("M").Value = final_mortality
    TBL.Fields("M_se_1st").Value = m1st
    TBL.Fields("M_se_2nd").Value = m2nd
    
    TBL.Fields("life_span").Value = lspan
    TBL.Fields("life_span_1st").Value = lspan_r1
    TBL.Fields("life_span_2nd").Value = lspan_r2
        
    
    TBL.Fields("generation_time").Value = GENTIME
    TBL.Fields("gen_time_1st").Value = gentime_r1
    TBL.Fields("gen_time_2nd").Value = gentime_r2
    
    
    
    TBL.Fields("tm").Value = gtime
    TBL.Fields("tm_1st").Value = gtime_r1
    TBL.Fields("tm_2nd").Value = gtime_r2
    
    
    
    TBL.Fields("Lm").Value = vlmaturity
    TBL.Fields("Lm_se_1st").Value = lm_1
    TBL.Fields("Lm_se_2nd").Value = lm_2
    TBL.Fields("Lm_type").Value = maturity_lt
    
    TBL.Fields("Lopt").Value = lmaxyield
    TBL.Fields("Lopt_se_1st").Value = lmaxyield_range1
    TBL.Fields("Lopt_se_2nd").Value = lmaxyield_range2
    TBL.Fields("Lopt_type").Value = yield_lt
    TBL.Fields("Lopt_text").Value = lmaxyield_est
        
        
    TBL.Fields("a").Value = var_a
    TBL.Fields("b").Value = var_b
    TBL.Fields("W").Value = finalval
    TBL.Fields("W_type").Value = finalval_type
    TBL.Fields("LW_length").Value = variable_length2
    TBL.Fields("LW_length_type").Value = variable_type3
    
    
    TBL.Fields("nitrogen").Value = nitrogen
    TBL.Fields("protein").Value = protein
    TBL.Fields("NitrogenProtein_weight").Value = np_weight
    
    TBL.Fields("reproductive_guild").Value = rg
    
    
    TBL.Fields("fecundity").Value = fecundity_v
    TBL.Fields("fecundity_1st").Value = fecundity_v1
    TBL.Fields("fecundity_2nd").Value = fecundity_v2
    'TBL.Fields("fecundity_text").Value = fecundity_text
    
    
    
    TBL.Fields("Emsy").Value = vemsy
    TBL.Fields("Eopt").Value = veopt
    TBL.Fields("Fmsy").Value = vfmsy
    TBL.Fields("Fopt").Value = vfopt
    TBL.Fields("Lc").Value = vLc
    TBL.Fields("Lc_type").Value = Lc_lt
    TBL.Fields("E").Value = vE
    TBL.Fields("YR").Value = vYR
    
    TBL.Fields("resilience").Value = resiliency
    
    
    
    TBL.Fields("rm").Value = Round(vrm, 2)
    TBL.Fields("Lr").Value = vlr
    TBL.Fields("Lr_type").Value = lr_lt
    
    
    TBL.Fields("main_food").Value = mf
    TBL.Fields("trophic_level").Value = tl
        
     
     
     
     
    
    
    TBL.Fields("QB").Value = finalqb
    TBL.Fields("QB_text").Value = finalqb_text
    TBL.Fields("QB_winf").Value = vwinf
    TBL.Fields("QB_temp").Value = var_temp_qb
    TBL.Fields("QB_A").Value = vAfin
    
    
    
    
    TBL.Update
        
        
    End If
    TBL.MoveNext
Wend


'MsgBox TBL.RecordCount
'MsgBox i
TBL.Close
MDB.Close




End Sub

Private Sub Command5_Click()
'4

Dim MDB As Database
Dim TBL As Recordset

Set MDB = OpenDatabase(App.Path & "\keyfacts.mdb")
Set TBL = MDB.OpenRecordset("SELECT Matrix.*, SPECIES.FamCode, SPECIES.Length, SPECIES.LengthFemale, SPECIES.LTypeMaxM, SPECIES.LTypeMaxF " & _
"FROM Matrix LEFT JOIN SPECIES ON Matrix.SpecCode = SPECIES.SpecCode;", dbOpenDynaset)

Dim i As Long
i = 0
Dim X As Long
Dim SSS As Variant
Dim STR As String

TBL.MoveFirst


While Not TBL.EOF
    
    
'start initialize c5
    Variable_Length = Null
    variable_type1 = Null
    variable_infinity = Null
    variable_type2 = Null
    linf_r1 = Null
    Linf_r2 = Null
    finalk = Null
    variable_q = Null
    xto = Null
    var_temp = Null
    final_mortality = Null
    m1st = Null
    m2nd = Null
    lspan = Null
    GENTIME = Null
    gentime_r1 = Null
    gentime_r2 = Null
    gtime = Null
    gtime_r1 = Null
    gtime_r2 = Null
    vlmaturity = Null
    lm_1 = Null
    lm_2 = Null
    maturity_lt = Null
    lmaxyield = Null
    lmaxyield_range1 = Null
    lmaxyield_range2 = Null
    yield_lt = Null
    lmaxyield_est = Null
    var_a = Null
    var_b = Null
    finalval = Null
    finalval_type = Null
    variable_length2 = Null
    variable_type3 = Null
    nitrogen = Null
    protein = Null
    np_weight = Null
    rg = Null
    fecundity_v = Null
    fecundity_v1 = Null
    fecundity_v2 = Null
    vemsy = Null
    veopt = Null
    vfmsy = Null
    vfopt = Null
    vLc = Null
    Lc_lt = Null
    vE = Null
    vYR = Null
    resiliency = Null
    vrm = Null
    vlr = Null
    lr_lt = Null
    mf = Null
    tl = Null
    finalqb = Null
    finalqb_text = Null
    vwinf = Null
    var_temp_qb = Null
    vAfin = Null

'end initialize
    
    
    If TBL!kf_type = 4 Then
    'If TBL!Speccode = 2 Or TBL!Speccode = 2 Then
        i = i + 1
        X = TBL!SpecCode
        var_tmax = TBL!tm_for_KF
        
    STR = "SELECT STOCKS.Stockcode From stocks " & _
    "where stocks.SpecCode=" & X & "and stocks.level = 'species in general'"
    Set getstockcode = MDB.OpenRecordset(STR, dbOpenDynaset)
    If getstockcode.RecordCount <> 0 Then
        vstockcode = getstockcode!StockCode
    Else
        vstockcode = 0
    End If
        
        
        
        
        
    '############################################################################
    '####### start main #########################################################
    '############################################################################
    var_temp = 0
        
    
    linf_r1 = Null
    Linf_r2 = Null
    
    '############################################################################
    'start lmax for KF_type = 1
    Variable_Length = 1
    If TBL!length <> "" And Not IsNull(TBL!length) Then
        Variable_Length = TBL!length
    Else
        If TBL!lengthfemale <> "" And Not IsNull(TBL!lengthfemale) Then
            Variable_Length = TBL!lengthfemale
        Else
            Variable_Length = 1
        End If
    End If
    If TBL!ltypemaxm <> "" And Not IsNull(TBL!ltypemaxm) Then
        variable_type1 = TBL!ltypemaxm
    Else
        variable_type1 = ""
    End If
    'end lmax for KF_type = 1
    
        

      
    
    
    
    '############################################################################
    
    
    '############################################################################
    'start linf for KF_type = 3
    
    
        If TBL!length <> "" And Not IsNull(TBL!length) Then
            variable_lmax = TBL!length
            variable_infinity = 10 ^ (0.044 + 0.9841 * (Log(TBL!length) / Log(10)))
            
            
            
    'If X = 502 Then
    'MsgBox TBL!length
    'MsgBox "aa " & variable_infinity
    'End If
    
        Else
            If TBL!lengthfemale <> "" And Not IsNull(TBL!lengthfemale) Then
                variable_lmax = TBL!lengthfemale
                variable_infinity = 10 ^ (0.044 + 0.9841 * (Log(TBL!lengthfemale) / Log(10)))
    'If X = 502 Then
    'MsgBox "bb " & variable_infinity
    'End If
            
            Else
                variable_lmax = 1
                variable_infinity = 10 ^ (0.044 + 0.9841 * (Log(1) / Log(10)))
    
    'If X = 502 Then
    'MsgBox "cc " & variable_infinity
    'End If
            
            End If
        End If
   
    variable_infinity = Round(variable_infinity, 1)
    'If X = 502 Then
    'MsgBox variable_infinity
    'End If
    
    vlmaturity = Round(10 ^ (0.898 * (Log(variable_infinity) / Log(10)) - 0.0781), 1)

    mat_r1 = Round(10 ^ (0.898 * (Log(variable_infinity) / Log(10)) - 0.0781 - 0.127), 1)
    mat_r2 = Round(10 ^ (0.898 * (Log(variable_infinity) / Log(10)) - 0.0781 + 0.127), 1)

      

    'If GetQprime.RecordCount = 0 Then
    If Not IsNull(TBL!length) Then
    
    linf_r1 = Round(10 ^ (0.044 + 0.9841 * (Log(TBL!length) / Log(10)) - 0.074), 1)
    Linf_r2 = Round(10 ^ (0.044 + 0.9841 * (Log(TBL!length) / Log(10)) + 0.074), 1)
    
    End If
    'End If
    
    
    
    'end linf for KF_type = 1
    '############################################################################
    
    
    '############################################################################
    'start K for KF_type = 4
    
    
    
'xxxk = -Log(1 - vlmaturity / variable_infinity) / (var_tm)
'xto = -1 * (10 ^ (-0.3922 - 0.2752 * (Log(variable_infinity) / Log(10)) - 1.038 * (Log(xxxk) / Log(10))))
'xto = Round(xto, 2)
'finalk = -Log(1 - vlmaturity / variable_infinity) / (var_tm - xto)
'finalk = Round(finalk, 2)
    
xxxk = 3 / var_tmax
xto = -1 * (10 ^ (-0.3922 - 0.2752 * (Log(variable_infinity) / Log(10)) - 1.038 * (Log(xxxk) / Log(10))))
finalk = 3 / (var_tmax - xto)
finalk = Round(finalk, 2)
     
    
    
    
    

    'end K for KF_type = 3
    '############################################################################

    
    
    
    
    '############################################################################
    'start var_temp
    
    If var_temp = 0 Then
        STR = "SELECT Avg(POPGROWTH.Temperature) AS AvgOfT " & _
        "From POPGROWTH " & _
        "WHERE (((POPGROWTH.SpecCode)=" & TBL!SpecCode & ") AND (Not (POPGROWTH.Temperature)=0 " & _
        "And (POPGROWTH.Temperature) Is Not Null))"
        Set avgt = MDB.OpenRecordset(STR, dbOpenDynaset)
        
        While Not avgt.EOF
            var_temp = avgt!avgoft
            avgt.MoveNext
        Wend
    End If
    
    
    
    
    If var_temp = 0 Then
    STR = "SELECT Avg(([tempmin]+[tempmax])/2) AS vAvg " & _
    "From STOCKS " & _
    "WHERE   (   ((stocks.tempmin)<>0 And (stocks.tempmin) Is Not Null) AND " & _
                "((stocks.tempmax)<>0 And (stocks.tempmax) Is Not Null) AND " & _
                "((stocks.Speccode)=" & TBL!SpecCode & ")    )"
    Set opt_stocks = MDB.OpenRecordset(STR, dbOpenDynaset)
    If opt_stocks!vAvg > 0 Then
        var_temp = opt_stocks!vAvg
    End If
    End If
    
    
If IsNull(var_temp) Then
    STR = "SELECT STOCKS.EnvTemp From STOCKS " & _
    "WHERE (((STOCKS.stockcode)=" & vstockcode & "));"
    Set opt_stocks = MDB.OpenRecordset(STR, dbOpenDynaset)
    
    While Not opt_stocks.EOF
        If opt_stocks!envtemp = "boreal" Then
            var_temp = 6
        ElseIf opt_stocks!envtemp = "deep-water" Then
            var_temp = 8
        ElseIf opt_stocks!envtemp = "high altitude" Then
            var_temp = 1
        ElseIf opt_stocks!envtemp = "polar" Then
            var_temp = 1
        ElseIf opt_stocks!envtemp = "subtropical" Then
            var_temp = 17
        ElseIf opt_stocks!envtemp = "temperate" Then
            var_temp = 10
        ElseIf opt_stocks!envtemp = "tropical" Then
            var_temp = 25
        End If
        opt_stocks.MoveNext
    Wend
End If
   
    
    'end var_temp
    '############################################################################
    
    
    '############################################################################
    'start variable_type2
    
    
    STR = "SELECT POPgrowth.Speccode,popgrowth.temperature From POPgrowth " & _
    "WHERE (((POPgrowth.Speccode)=" & X & "))"
    Set withgrowth = MDB.OpenRecordset(STR, dbOpenDynaset)
    
    'If withgrowth.RecordCount = 0 Then
    '    If variable_type1 <> "" Then
    '        variable_type2 = variable_type1
    '    Else
    '        variable_type2 = ""
    '    End If
    'Else
    '    variable_type2 = xxxVT2
    'End If
    'If withgrowth.RecordCount = 0 Then
    '    variable_type2 = variable_type1
    'Else
    '    variable_type2 = variable_type2
    'End If
    
    variable_type2 = variable_type1
    
    
    
    'end variable_type2
    '############################################################################
    
    
    '############################################################################
    'start M and others
    final_mortality = 10 ^ (0.333 - 0.246 * (Log(variable_infinity) / Log(10)) + 0.744 * (Log(finalk) / Log(10)) + 0.01 * var_temp)
    m1st = Null
    m2nd = Null
    
    If variable_type2 = "TL" Then
        m1st = 10 ^ (0.333 - 0.246 * (Log(variable_infinity) / Log(10)) + 0.744 * (Log(finalk) / Log(10)) + 0.01 * var_temp - 0.18)
        m2nd = 10 ^ (0.333 - 0.246 * (Log(variable_infinity) / Log(10)) + 0.744 * (Log(finalk) / Log(10)) + 0.01 * var_temp + 0.18)
    Else
        'm1st = ""
        'm2nd = ""
    End If
    'end M and others
    '############################################################################
    
    
    '############################################################################
    'start life span
    lspan = (3 / finalk) + xto
    
If Not IsNull(linf_r1) Then
k1 = Round((-Log(1 - vlmaturity / linf_r1) / (var_tm - xto)), 2)
k2 = Round((-Log(1 - vlmaturity / Linf_r2) / (var_tm - xto)), 2)
lspan_r1 = Round((3 / k1) + xto, 1)
lspan_r2 = Round((3 / k2) + xto, 1)
End If
    
    
    
    
    'end life span
    '############################################################################
    
    
    '############################################################################
    'start gentime
    
    
'STR = "SELECT popgrowthref,loo,k,type,lm,tlinfinity From POPGROWTH " & _
'"WHERE (((POPGROWTH.SpecCode)=" & X & ")) order by popgrowth.loo,popgrowth.type;"
'Set medgrowth = MDB.OpenRecordset(STR, dbOpenDynaset)
'If medgrowth.RecordCount = 0 Then
'Else
'    vlmaturity = 10 ^ (0.898 * (Log(variable_infinity) / Log(10)) - 0.0781)
'    If (xxxLoo) Then '<!--- meaning with growth, median Qprime --->
'        qtest1 = variable_infinity * (3 / (3 + final_mortality / finalk))
'        qtest2 = 10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742)
'        If qtest1 < qtest2 Then
'            varlopt = 10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742)
'        Else
'            varlopt = variable_infinity * (3 / (3 + final_mortality / finalk))
'        End If
'    Else
'        varlopt = 10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742)
'    End If
'    evlmaturity = 10 ^ (0.898 * (Log(variable_infinity) / Log(10)) - 0.0781)
'    If evlmaturity >= varlopt Then
'        lm100 = evlmaturity + (variable_infinity - evlmaturity) / 4
'        gentime = xto + (-1 * (Log(1 - Val(lm100) / Val(variable_infinity)) / Val(finalk)))
'    Else
'        gentime = xto + (-1 * (Log(1 - Val(varlopt) / Val(variable_infinity)) / Val(finalk)))
'    End If
'End If
        
        
        
varlopt = Round(10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742), 1)

evlmaturity = Round(10 ^ (0.898 * (Log(variable_infinity) / Log(10)) - 0.0781), 1)
If evlmaturity >= varlopt Then
    lm100 = evlmaturity + (variable_infinity - evlmaturity) / 4
    GENTIME = xto + (-1 * (Log(1 - lm100 / variable_infinity) / finalk))
    
    'If X = 502 Then
    '    MsgBox "111222"
    'End If
Else
       
lmaxyield_range1 = Round(10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742 - 0.073), 1)
lmaxyield_range2 = Round(10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742 + 0.073), 1)

       
       
    GENTIME = xto + (-1 * (Log(1 - varlopt / variable_infinity) / finalk))
    gentime_r1 = xto + (-1 * (Log(1 - lmaxyield_range1 / variable_infinity) / finalk))
    gentime_r2 = xto + (-1 * (Log(1 - lmaxyield_range2 / variable_infinity) / finalk))
    
    'If X = 502 Then
    '    MsgBox "333444"
    'End If

End If

        
        
        
        
        
        
    
    'end gentime
    '############################################################################
    
    '############################################################################
    'start tm
    
    'for kf_type=1
    'gtime = Round(xto, 2) + (-1 * (Log(1 - Val(Round(vlmaturity, 1)) / Val(Round(variable_infinity, 2))) / Val(Round(finalk, 2))))
    
    
    gtime = Round(xto + (-1 * (Log(1 - vlmaturity / variable_infinity) / finalk)), 1)
    
    'If X = 845 Then
    '    MsgBox xto & " " & vlmaturity & " " & variable_infinity & " " & finalk
    'End If
    
    If Not IsNull(linf_r1) Then
    gtime_r1 = Round(xto + (-1 * (Log(1 - mat_r1 / linf_r1) / finalk)), 1)
    gtime_r2 = Round(xto + (-1 * (Log(1 - mat_r2 / Linf_r2) / finalk)), 1)
    End If
    


    
    
    
    
    'If TBL!speccode = 2 Then
    'MsgBox Round(xto, 2)
    'MsgBox Round(vlmaturity, 1)
    'MsgBox Round(variable_infinity, 2)
    'MsgBox Round(finalk, 2)
    'End If
    
    'end tm
    '############################################################################
        
        
        
    '############################################################################
    'start Lm se
    lm_1 = 10 ^ (0.898 * (Log(variable_infinity) / Log(10)) - 0.0781 - 0.127)
    lm_2 = 10 ^ (0.898 * (Log(variable_infinity) / Log(10)) - 0.0781 + 0.127)
    maturity_lt = variable_type2
    'end Lm se
    '############################################################################
    
        
    '############################################################################
    'start lopt
    
    If xxxLoo <> 0 Then '<!--- meaning with growth, median Qprime --->
        qtest1 = variable_infinity * (3 / (3 + final_mortality / finalk))
        qtest2 = 10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742)
        If qtest1 < qtest2 Then
            lmaxyield = 10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742)
            'MsgBox "2"
        Else
            lmaxyield = variable_infinity * (3 / (3 + Round(final_mortality, 2) / Round(finalk, 2)))
            'MsgBox "3"
            'MsgBox lmaxyield
            'MsgBox variable_infinity
            'MsgBox final_mortality
            'MsgBox finalk
        End If
    Else
        lmaxyield = 10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742)
        'MsgBox "4"
    End If
    
    
    'this is from kf_type=1
    'If xxxLoo = 0 Then '<!--- meaning w/o growth, median Qprime --->
    '    lmaxyield_est = "Estimated from Linf."
    'Else
    
        'ito ay parang pang kf_type=1
        'If qtest1 < qtest2 Then
        '    lmaxyield_range1 = 10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742 - 0.073)
        '    lmaxyield_range2 = 10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742 + 0.073)
        '    lmaxyield_est = "Estimated from Linf."
        'Else
        '    lmaxyield_range1 = Null
        '    lmaxyield_range2 = Null
        '    lmaxyield_est = "Estimated from Linf., K and M."
        'End If
    'End If
    
    
    
    
    
    
        
    yield_lt = variable_type2
    
    'end lopt
    '############################################################################
    
    
    '############################################################################
    'start of l-w
    variable_length2 = Round(variable_infinity, 1)


    'start <!--- get the median of LW --->
    STR = "SELECT  POPLW.SpecCode,POPLW.LengthMin,POPLW.LengthMax,POPLW.Number,POPLW.Sex, POPLW.a, POPLW.b, COUNTREF.paese, " & _
        "poplw.autoctr, poplw.locality,poplw.type, poplw.a , poplw.b FROM COUNTREF INNER JOIN POPLW ON COUNTREF.C_Code = POPLW.C_Code " & _
    "WHERE (((POPLW.SpecCode)=" & TBL!SpecCode & "))        order by poplw.b"
    Set medlwb = MDB.OpenRecordset(STR, dbOpenDynaset)
    
    If medlwb.RecordCount <> 0 Then
        xLW = (medlwb.RecordCount / 2) + 0.5
        xLW = Round(xLW)
        
        'If X = 68 Then
        '    MsgBox xLW
        'End If
        
        ii = 0
        While Not medlwb.EOF
            ii = ii + 1
            If ii = xLW Then
                v_a = medlwb!a
                v_b = medlwb!b
                lwtype = medlwb!Type
            End If
            medlwb.MoveNext
        Wend
    End If
    'end <!--- get the median of LW --->



'<!--- start of lw --->
If medlwb.RecordCount <> 0 Then
    'start parang hindi dada-anan ito
    'if #parameterexists(variable_length2)# is "No">
    '    variable_length2 = #numberformat(length,"9999.9")#>
    'end if
    'end parang hindi dada-anan ito
    
    
    'if #parameterexists(var_a)# is "No">
    var_a = v_a
    'end if
    'if #parameterexists(var_b)# is "No">
    var_b = v_b
    'end if
    finalval = ((variable_length2 ^ Round(var_b, 3)) * Round(var_a, 4))
    'finalval = ((46 ^ 3.13) * 0.0054)
    
    'If X = 68 Then
    '    MsgBox variable_length2 & " " & var_b & " " & var_a & " = " & finalval
    'End If
    
    variable_type3 = lwtype
    eli01 = Len(Trim(Round(finalval, 1))) * 2


    'If TBL!speccode = 2 Then
    '    MsgBox var_a
    '    MsgBox var_b
    '    MsgBox finalval
    'End If

    
    np_weight = finalval
    
    If finalval > 200 Then
        If finalval > 20000 Then
            finalval = Round(finalval / 1000, 1)
            finalval_type = "kg"
        Else
            finalval = Round(finalval)
            finalval_type = "g"
        End If
    Else
        finalval = Round(finalval, 1)
        finalval_type = "g"
    End If
    
       
    'W = finalval
    'a = Var_a
    'b = Var_b
   
    'Nitrogen & protein:
    
    '<!--- start new weight --->
    If finalval <> 0 Then
        'NP_weight = finalval   .... transferred up to preserve the decimal places..
        
        'If Trim(finalval_type) = "kg" Then
        '    NP_weight = NP_weight * 1000
        'End If
        
        
        
        nitrogen = 10 ^ (1.03 * (Log(np_weight) / Log(10)) - 1.65)
        nitrogen = Round(nitrogen, 1)
        protein = Round(6.25 * nitrogen, 1)
        np_weight = Round(np_weight)
        
        
        
    Else
        np_weight = Null
        nitrogen = Null
        protein = Null
    End If
    '<!--- end new weight --->
    
Else
    '<!--- don't show --->
End If

'<!--- end of lw --->
    
    
    'end of l-w
    '############################################################################
    
    
    '############################################################################
    'start reproductive_guild
    rg = ""
    If withgrowth.RecordCount <> 0 Then
    
    'MsgBox withgrowth.RecordCount
    'MsgBox "riz = " & vstockcode
    
    If withgrowth.RecordCount <> 0 Then
    
    
    
    STR = "SELECT REPRODUC.RepGuild1, REPRODUC.RepGuild2 From REPRODUC " & _
    "WHERE (((REPRODUC.StockCode)=" & vstockcode & "))"
    Set getguild = MDB.OpenRecordset(STR, dbOpenDynaset)
    rg = ""
    If getguild.RecordCount <> 0 Then
        If Trim(getguild!repguild1) = "" Then
            rg = " "
        Else
            rg = getguild!repguild1
        End If
        rg = rg + ": "
        If Trim(getguild.repguild2) = "" Then
            rg = rg & " "
        Else
            rg = rg & getguild!repguild2
        End If
        rg = Trim(rg)
    End If
    End If
    End If
    'end reproductive_guild
    '############################################################################
    

    
    
    
    '############################################################################
    'start fecundity
    
    fecundity_v = Null
    fecundity_v1 = Null
    fecundity_v2 = Null
    fecundity_text = Null
    
    
    
    If withgrowth.RecordCount <> 0 And vstockcode <> 0 Then
    
    'If X = 69 Then
    'MsgBox X
    'MsgBox Vstockcode
    'End If

STR = "SELECT Min(SPAWNING.FecundityMin) AS MinOfFecundityMin From SPAWNING " & _
"GROUP BY SPAWNING.StockCode, SPAWNING.SpecCode " & _
"HAVING ((Not (Min(SPAWNING.FecundityMin))=0) AND ((SPAWNING.StockCode)=" & vstockcode & ") AND ((SPAWNING.SpecCode)=" & X & ")) " & _
"ORDER BY Min(SPAWNING.FecundityMin)"
Set getmin = MDB.OpenRecordset(STR, dbOpenDynaset)

STR = "SELECT Max(SPAWNING.FecundityMax) AS MaxOfFecundityMax From SPAWNING GROUP BY SPAWNING.StockCode, SPAWNING.SpecCode " & _
"HAVING (((SPAWNING.StockCode)=" & vstockcode & ") AND ((SPAWNING.SpecCode)=" & X & ") AND (Not (Max(SPAWNING.FecundityMax))=0))"
Set getmax = MDB.OpenRecordset(STR, dbOpenDynaset)
    
    
If getmin.RecordCount <> 0 Or getmax.RecordCount <> 0 Then
    
    y01 = 0
    y02 = 0
    fecundity_v1 = ""
    fecundity_v2 = ""
    
    
    If getmin.RecordCount <> 0 And getmax.RecordCount <> 0 Then
    If getmin!minoffecunditymin > 0 And getmax!maxoffecunditymax > 0 Then
    y01 = Log(getmin.minoffecunditymin) / Log(10)
    y02 = Log(getmax.maxoffecunditymax) / Log(10)
    y01 = y01 + y02
    y01 = y01 / 2
    fecundity_v = Round(10 ^ y01)
    End If
    End If
    
    If getmin.RecordCount <> 0 Then
        If getmin!minoffecunditymin = "" Then
            fecundity_v1 = "no value (min.)"
        Else
            fecundity_v1 = Round(getmin!minoffecunditymin)
        End If
    Else
        fecundity_v1 = "no record (min.)"
    End If
    
    
    If getmax.RecordCount <> 0 Then
        If getmax!maxoffecunditymax = "" Then
            fecundity_v2 = "no value (max.)"
        Else
            fecundity_v2 = Round(getmax!maxoffecunditymax)
        End If
    Else
        fecundity_v2 = "no record (max.)"
    End If
    
    
    
    If getmin.RecordCount <> 0 And getmax.RecordCount <> 0 Then
        fecundity_text = "Estimated as geometric mean."
    End If


End If
    End If
    'end fecundity
    '############################################################################
    
    
    
'############################################################################
'start yrecruit

'<!--- start yrecruit --->

    vemsy = Null
    veopt = Null
    vfmsy = Null
    vfmsy_rm = Null
    vfopt = Null
    vLc = Null
    Lc_lt = Null
    vE = Null
    vYR = Null





vE = 0.5
vLc = Round(0.4 * variable_infinity, 1)
If final_mortality > 0 Then
    
    vU = 1 - vLc / Round(variable_infinity, 2)
    MK = Round(final_mortality, 2) / Round(finalk, 2)
    
    vx1 = 0
    vx2 = 0
    firstloop = "Y"
    
    oldy = 0
    vlope = 0
    vemsy = 0
    veopt = 0
    vfmsy = 0
    vfmsy_rm = 0
    vfopt = 0
    
    
    While vx1 <= 1
        
        vx1 = vx2
        vx2 = vx2 + 0.001

        vm1 = (1 - vx1) / MK
        vy1 = vx1 * (vU ^ MK) * (1 - ((3 * vU) / (1 + vm1)) + ((3 * vU ^ 2) / (1 + 2 * vm1)) - ((vU ^ 3) / (1 + 3 * vm1)))
        vm2 = (1 - vx2) / MK
        vy2 = vx2 * (vU ^ MK) * (1 - ((3 * vU) / (1 + vm2)) + ((3 * vU ^ 2) / (1 + 2 * vm2)) - ((vU ^ 3) / (1 + 3 * vm2)))
        vslope = (vy2 - vy1) / (vx2 - vx1)
        
        If oldy <> 0 Then
            If vy1 >= oldy Then
            Else
                vemsy = Round(vx1 - 0.001, 2)
                '<cfbreak>
                GoTo EliGo
            End If
        End If
        
        
        oldy = vy1
        
        If firstloop = "Y" Then
            firstvalue = (vy2 - vy1) / (vx2 - vx1)
            firstloop = "N"
        End If

        
    
        If veopt = 0 Then
        If Round(vslope, 3) = Round(firstvalue / 10, 3) Then
            veopt = vx1
        End If
        End If
        
        
    Wend
EliGo:
    '<!--- end get e  --->
    

    vm = (1 - vE) / MK
    lijosh = vE * (vU ^ MK) * (1 - ((3 * vU) / (1 + vm)) + ((3 * vU ^ 2) / (1 + 2 * vm)) - ((vU ^ 3) / (1 + 3 * vm)))
    '<input type="text" name="vYR" value=round(lijosh,'9.9999')#" size="6" onFocus="noedit(2)"  align="right">
    vYR = Round(lijosh, 4)

Else
    '<input type="text" name="vYR" value="" size="6" onFocus="noedit(2)"  align="right">
    vYR = Null
End If

    vLc = Round(vLc, 1)
    Lc_lt = Trim(variable_type2)
    e = vE = Round(vE, 2)
    
    If final_mortality > 0 Then
    
        veopt = Round(veopt, 2)
        vemsy = Round(vemsy, 2)
        'vfmsy = Round(final_mortality * vEmsy / (1 - vEmsy), 2)
        vfmsy = final_mortality * vemsy / (1 - vemsy)
        vfmsy_rm = final_mortality * vemsy / (1 - vemsy)
        'vFopt = Round(final_mortality * vEopt / (1 - vEopt), 2)
        vfopt = final_mortality * veopt / (1 - veopt)
        
        
        
        
        
        
        veopt = Round(veopt, 2)
        vemsy = Round(vemsy, 2)
        vfmsy = Round(vfmsy, 2)
        vfopt = Round(vfopt, 2)
    Else
        vemsy = Null
        veopt = Null
        vfmsy = Null
        vfmsy_rm = Null
        vfopt = Null
    End If





'<!--- end yrecruit   width  msy --->


'end yrecruit
'############################################################################
    
    
    '############################################################################
    'start resiliency
    resiliency = Null
    
    
If finalk <= 0.05 Or var_tmax > 30 Then
    resiliency = "Very low; decline threshold 0.70"
ElseIf finalk <= 0.15 Or var_tmax >= 11 Then
    resiliency = "Low; decline threshold 0.85"
ElseIf finalk <= 0.3 Or var_tmax >= 4 Then
    resiliency = "Medium; decline threshold 0.95"
ElseIf finalk > 0.3 Or var_tmax < 4 Then
    resiliency = "High; decline threshold 0.99"
Else
    resiliency = "Please enter values for K, tmax."
End If
    
    
    
    
    
    
    
    'If Round(finalk, 2) <= 0.05 Then
    '    resiliency = "Very low; decline threshold 0.70"
    'ElseIf Round(finalk, 2) <= 0.15 Then
    '    resiliency = "Low; decline threshold 0.85"
    'ElseIf Round(finalk, 2) <= 0.3 Then
    '    resiliency = "Medium; decline threshold 0.95"
    'ElseIf Round(finalk, 2) > 0.3 Then
    '    resiliency = "High; decline threshold 0.99"
    'Else
    '    resiliency = "Please enter values for K."
    'End If
    
    
    'end resiliency
    '############################################################################
    
    
    
    
    
    
    '############################################################################
    'start rm
    vlr = Round(0.4 * variable_infinity, 1)
    vfmsy = Round(vfmsy, 2)
    vrm = Round(2 * vfmsy, 2)
    vypdt = Log(2) / vrm
    lr_lt = Trim(variable_type2)
    
    'If X = 2 Then
    '    MsgBox vfmsy
    '    MsgBox vrm
    'End If
    
    'end rm
    '############################################################################
    
    
    
    '############################################################################
    'start eco
    STR = "SELECT  ECOLOGY.StockCode,ECOLOGY.DietTroph,ECOLOGY.DietSeTroph,ECOLOGY.FoodTroph, " & _
    "ECOLOGY.FoodSeTroph,0 as EcoTroph,0 as EcoSeTroph,'' as mainfood,ECOLOGY.herbivory2 " & _
    "From ECOLOGY WHERE (((ECOLOGY.StockCode)=" & vstockcode & "))"
    Set witheco = MDB.OpenRecordset(STR, dbOpenDynaset)
    
mf = Null
tl = Null
If witheco.RecordCount <> 0 Then
    If witheco!mainfood <> "" Then
        mf = witheco!mainfood
    End If
    If witheco!herbivory2 <> "" Then
        mf = mf & " " & witheco!herbivory2
    End If
    
    
    'Trophic level:
    tl = Null
    If witheco!dietTroph <> "" Then
        tl = Round(witheco!dietTroph, 1)
        If witheco!DietSeTroph <> 0 Then
            tl = tl & "&nbsp;&nbsp;&nbsp;"
            tl = tl & "+/- s.e. " & Round(witheco!DietSeTroph, 2)
        End If
        tl = tl & " Estimated from diet data."
    Else
        If witheco!foodTroph <> "" Then
            tl = tl & " " & Round(witheco!foodTroph, 1)
            If witheco!FoodSeTroph <> 0 Then
                tl = tl & "+/- s.e. " & Round(witheco!FoodSeTroph, 2)
            End If
            tl = tl & " Estimated from food data."
        Else
            If witheco!EcoTroph <> "" Then
                tl = tl & " " & Round(witheco!EcoTroph, 1)
                If witheco!EcoSeTroph <> 0 Then
                    tl = tl & "+/- s.e. " & Round(witheco!EcoSeTroph, 2)
                End If
                tl = tl & " Estimated from Ecopath model."
            End If
        End If
    End If
End If


'<!---start of food consumption  </cfif --->

    
    'end eco
    '############################################################################
    
    
    
    
    
    
    
    
    '############################################################################
    'start qb
    
    finalqb = Null
    finalqb_text = Null
    vwinf = Null
    var_temp_qb = Null
    vAfin = Null
    
    
    
    
    
'<!--- start food consumption --->
contqb = "Y"
'<!--- start get median popqb --->
STR = "SELECT popqb.popqb FROM popqb WHERE (((popqb.speccode)=" & X & ")) ORDER BY popqb.popqb"
Set qbmedian = MDB.OpenRecordset(STR, dbOpenDynaset)


'start of old
'vmedian = (qbmedian.RecordCount / 2) + 0.5
'vmedian = Round(vmedian)
'end of old


vMedian = (qbmedian.RecordCount / 2) + 0.5
int_part = Int(vMedian)
str_conv = "" & vMedian

If Len(str_conv) = 1 Then
    vMedian = Round(vMedian)
Else

    If Mid(str_conv, Len(str_conv) - 1, 2) = ".5" Then
        vMedian = int_part + 1
    Else
        vMedian = Round(vMedian)
    End If
End If




vpopqb = 0


'If X = 69 Then
'    MsgBox qbmedian.RecordCount
'    MsgBox (qbmedian.RecordCount / 2) + 0.5
'    MsgBox Int((qbmedian.RecordCount / 2) + 0.5)
'    MsgBox Round(vmedian)
'    MsgBox Round(vmedian, 1)
'End If


ii = 0
While Not qbmedian.EOF

    ii = ii + 1
    If ii = vMedian Then
        vpopqb = qbmedian!popqb
    End If
    qbmedian.MoveNext
Wend


'<!--- end get median popqb --->

If vpopqb > 0 Then
    explain = "with popqb record"
    contqb = "N"
Else
    explain = "no popqb record"
    
    '<!---#############################################################################################--->
    '<!---### start of no popqb #######################################################################--->
    '<!---#############################################################################################--->
    If medlwb.RecordCount# > 0 Then
        explain = explain & "; with lw rel"
    Else
        explain = explain & "; no lw rel"
    End If

    '<!--- start of A --->
    STR = "SELECT Swimming.AspectRatio as aspectratio From Swimming WHERE (((Swimming.SpecCode)=" & X & "));"
    Set getAR = MDB.OpenRecordset(STR, dbOpenDynaset)
    vAfin = 0
    If getAR.RecordCount > 0 Then
       While Not getAR.EOF
            If getAR!aspectratio > 0 Then
                vAfin = getAR!aspectratio
                explain = explain & "; with aspect ratio"
            Else
                vAfin = 0
                explain = explain & "; w/o aspect ratio"
            End If
            getAR.MoveNext
        Wend
    Else
        explain = explain & "; w/o aspect ratio"
    End If
    '<!--- end of A --->
    
    
    '<!--- start of h and d --->
    SelectHD = "N"
    '<!---
    '    go ecology feeding type
    '    if none troph diettroph
    '        if none troph foodtroph items
    '            if none troph ecotroph
    '--->


    STR = "SELECT ECOLOGY.Herbivory2, 0 as EcoTroph, ECOLOGY.DietTroph, ECOLOGY.FoodTroph From ECOLOGY " & _
    "WHERE (((ECOLOGY.StockCode)=" & vstockcode & "))"
    Set getft = MDB.OpenRecordset(STR, dbOpenDynaset)

    
    
    If getft.RecordCount <> 0 Then
    If Trim(getft!herbivory2) <> "" Then
    
        If Trim(getft!herbivory2) = "mainly animals (troph. 2.8 and up)" Then
            vh = 0
            vd = 0
        ElseIf Trim(getft!herbivory2) = "mainly animals (troph. 2.8 up)" Then
            vh = 0
            vd = 0
        ElseIf Trim(getft!herbivory2) = "mainly plants/detritus (troph. 2-2.19)" Then
            vh = 0
            vd = 1
        ElseIf Trim(getft.herbivory2) = "plants/detritus+animals (troph. 2.2-2.79)" Then
            vh = 1
            vd = 0
        End If
        explain = explain & "; w/ feeding type"
        WithFT = 1
    
    Else
    
    
        explain = explain & "; w/o feeding type"
        WithFT = 0
        If (getft!dietTroph) > 0 Then
            If getft!dietTroph >= 2 And getft!dietTroph <= 2.19 Then
                vh = 0
                vd = 1
            ElseIf getft!dietTroph >= 2.2 And getft!dietTroph <= 2.79 Then
                vh = 1
                vd = 0
            ElseIf getft!dietTroph >= 2.8 Then
                vh = 0
                vd = 0
            End If
            explain = explain & "; from diettroph"
        Else
            If getft!foodTroph > 0 Then
                If getft!foodTroph >= 2 And getft.foodTroph <= 2.19 Then
                    vh = 0
                    vd = 1
                ElseIf getft.foodTroph >= 2.2 And getft.foodTroph <= 2.79 Then
                    vh = 1
                    vd = 0
                ElseIf getft!foodTroph >= 2.8 Then
                    vh = 0
                    vd = 0
                End If
                explain = explain & "; from foodtroph"
            Else
                If getft!EcoTroph > 0 Then
                    If getft!EcoTroph >= 2 And getft!EcoTroph <= 2.19 Then
                        vh = 0
                        vd = 1
                    ElseIf getft!EcoTroph >= 2.2 And getft!EcoTroph <= 2.79 Then
                        vh = 1
                        vd = 0
                    ElseIf getft!EcoTroph >= 2.8 Then
                        vh = 0
                        vd = 0
                    End If
                    explain = explain & "; from ecotroph"
                Else
                    explain = explain & "; no diet,food,eco trophs; select h,d"
                    '<!--- blank means yes
                    'contqb = "Y">
                    '--->
                    SelectHD = "Y"
                    '<!---
                    'contqb = "N">
                    'cont. with the search
                    'now the genus of median of diet,food,eco
                    '--->
                End If
            End If
        End If
    End If
    End If
    '<!--- end of h and d --->

'<!---#############################################################################################--->
'<!---### end of no popqb #########################################################################--->
'<!---#############################################################################################--->
End If
'<!--- end food consumption --->









'<!---start of food consumption  </cfif --->

'<td align="left">Food consumption (Q/B):</td>
If vpopqb > 0 Then
    finalqb = Round(vpopqb, 2)
    finalqb_text = "times the body weight per year"
Else
    '<!---#############################################################################################--->
    '<!---#############################################################################################--->
    '<!---#############################################################################################--->
    If contqb = "Y" Then
        '<!--- start for winf --->
        If medlwb.RecordCount > 0 Then
            'orig
            'vwinf = finalval
            
            'use this instead as the finalval is sometimes in kg
            vwinf = np_weight
            'If X = 68 Then
            '    MsgBox "111"
            'End If
            
            whereWinf = "lw"
        Else
            vwinf = 0.01 * variable_infinity ^ 3
            'If X = 68 Then
            '    MsgBox "222"
            'End If
            
            whereWinf = "DP"
        End If
        '<!--- end for winf --->
        If SelectHD = "Y" Then
            vh = 3
            vd = 3
        End If
        'start vb comment
        '<input type="hidden" name="vh" value="#vh#" size="1" onFocus="noedit(2)"  >
        '<input type="hidden" name="vd" value="#vd#" size="1" onFocus="noedit(2)"  >
        'end vb comment
        
        If vAfin > 0 Then
            xyz = vAfin
        Else
            xyz = 1.32
        End If
        If vh = 3 Then
            '<!---w/o h, d --->
            elix = 10 ^ (7.964 - 0.204 * Log(vwinf) / Log(10) - 1.965 * (1000 / (var_temp + 273.15)) + 0.083 * xyz + 0.532 * 1 + 0.398 * 0)
            elix2 = 10 ^ (7.964 - 0.204 * Log(vwinf) / Log(10) - 1.965 * (1000 / (var_temp + 273.15)) + 0.083 * xyz + 0.532 * 0 + 0.398 * 0)
            elix = (elix2 + elix) / 2
        Else
            '<!---with h, d--->
            elix = 10 ^ (7.964 - 0.204 * Log(vwinf) / Log(10) - 1.965 * (1000 / (var_temp + 273.15)) + 0.083 * xyz + 0.532 * vh + 0.398 * vd)
        End If
        
        
        If whereWinf = "none" Then
            finalqb = Null
        Else
            finalqb = Round(elix, 1)
        End If
        finalqb_text = "times the body weight per year"
        
        
        'Enter Winf, temperature, aspect ratio (A), and food type to estimate Q/B
        If whereWinf = "none" Then
            'Winf =
            vwinf = Null
            'If X = 68 Then
            '    MsgBox "333"
            'End If
            
            'Temp. =
            var_temp_qb = Round(var_temp, 1)
        Else
            'Winf =
            vwinf = Round(vwinf, 1)
            
            'If X = 68 Then
            '    MsgBox "444"
            'End If
            
            'Temp. =
            var_temp_qb = Round(var_temp, 1)
        End If
        
        If vAfin > 0 Then
            'A =
            vAfin = Round(vAfin, 2)
        Else
            vAfin = 1.32
            'A =
            vAfin = Round(vAfin, 2)
            
            '<!---eeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeee--->
            
            'start vb comment
            '<td colspan="8"><img src="../jpgs/Tails.gif" height=29 border=0 alt=""></td>
            '<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
            '<td><input type="radio" name="eli"          onClick="finc('#contqb#',6.55)"></td>
            '<td><input type="radio" name="eli"          onClick="finc('#contqb#',1.9)"></td>
            '<td><input type="radio"  checked name="eli" onClick="finc('#contqb#',1.32)"></td>
            '<td><input type="radio" name="eli"          onClick="finc('#contqb#',1.32)"></td>
            '<td><input type="radio" name="eli"          onClick="finc('#contqb#',1.63)"></td>
            'end vb comment
            
        End If
        If SelectHD = "Y" Then
            'start vb comment
            '<input type="hidden" name="omni" value="1">
            '<td align="center"><font size="2">Detrivore</font></td>
            '<td align="center"><font size="2">Herbivore</font></td>
            '<td align="center"><font size="2">Omnivore</font></td>
            '<td align="center"><font size="2">Carnivore</font></td>
            '<td align="center"><input type="radio" name="riz"           onClick="changeHD('#contqb#',0,1)"></td>
            '<td align="center"><input type="radio" name="riz"          onClick="changeHD('#contqb#',1,0)"></td>
            '<td align="center"><input type="radio" name="riz" checked  onClick="changeHD('#contqb#',3,3)"></td>
            '<td align="center"><input type="radio" name="riz"           onClick="changeHD('#contqb#',0,0)"></td>
            '</tr></table>
            'end vb comment
            
        Else
            'start vb comment
            '<input type="hidden" name="omni" value="0">
            '<td align="center"><font size="2">Detrivore</font></td>
            '<td align="center"><font size="2">Herbivore</font></td>
            '<td align="center"><font size="2">Omnivore</font></td>
            '<td align="center"><font size="2">Carnivore</font></td>
            '<td align="center"><input type="radio" name="riz" if #vh# is 0 and #vd# is 1>checkedend if onClick="changeHD('#contqb#',0,1)"></td>
            '<td align="center"><input type="radio" name="riz" if #vh# is 1 and #vd# is 0>checkedend if onClick="changeHD('#contqb#',1,0)"></td>
            '<td align="center"><input type="radio" name="riz"                                              onClick="changeHD('#contqb#',3,3)"></td>
            '<td align="center"><input type="radio" name="riz" if #vh# is 0 and #vd# is 0>checkedend if onClick="changeHD('#contqb#',0,0)"></td>
            'end vb comment
        End If
    End If '<!--- if #contqb# is "Y"> --->
    '<!---#############################################################################################--->
    '<!---#############################################################################################--->
    '<!---#############################################################################################--->
End If
'<!--- end of food consumption  </cfif estimate width --->


   
    
    'end qb
    '############################################################################
    
    
    
    
    
    
    
    
    
    
    
        
    '############################################################################
    '####### end main ###########################################################
    '############################################################################
                
        
    TBL.Edit
    TBL.Fields("lmax").Value = Variable_Length
    TBL.Fields("lmax_type").Value = variable_type1
    
    TBL.Fields("linf").Value = variable_infinity
    TBL.Fields("linf_type").Value = variable_type2
    
    TBL.Fields("linf_1st").Value = linf_r1
    TBL.Fields("linf_2nd").Value = Linf_r2
    
   
    
    TBL.Fields("K").Value = finalk
    TBL.Fields("PhiPrime").Value = variable_q
    TBL.Fields("to").Value = xto
    
    TBL.Fields("mean_temp").Value = var_temp
    
    TBL.Fields("M").Value = final_mortality
    TBL.Fields("M_se_1st").Value = m1st
    TBL.Fields("M_se_2nd").Value = m2nd
    
    TBL.Fields("life_span").Value = lspan
    'not needed here
    'TBL.Fields("life_span_1st").Value = lspan_r1
    'TBL.Fields("life_span_2nd").Value = lspan_r2
        
    
    TBL.Fields("generation_time").Value = GENTIME
    TBL.Fields("gen_time_1st").Value = gentime_r1
    TBL.Fields("gen_time_2nd").Value = gentime_r2
    
    
    
    TBL.Fields("tm").Value = gtime
    TBL.Fields("tm_1st").Value = gtime_r1
    TBL.Fields("tm_2nd").Value = gtime_r2
    
    
    
    TBL.Fields("Lm").Value = vlmaturity
    TBL.Fields("Lm_se_1st").Value = lm_1
    TBL.Fields("Lm_se_2nd").Value = lm_2
    TBL.Fields("Lm_type").Value = maturity_lt
    
    TBL.Fields("Lopt").Value = lmaxyield
    TBL.Fields("Lopt_se_1st").Value = lmaxyield_range1
    TBL.Fields("Lopt_se_2nd").Value = lmaxyield_range2
    TBL.Fields("Lopt_type").Value = yield_lt
    TBL.Fields("Lopt_text").Value = lmaxyield_est
        
        
    TBL.Fields("a").Value = var_a
    TBL.Fields("b").Value = var_b
    TBL.Fields("W").Value = finalval
    TBL.Fields("W_type").Value = finalval_type
    TBL.Fields("LW_length").Value = variable_length2
    TBL.Fields("LW_length_type").Value = variable_type3
    
    
    TBL.Fields("nitrogen").Value = nitrogen
    TBL.Fields("protein").Value = protein
    TBL.Fields("NitrogenProtein_weight").Value = np_weight
    
    TBL.Fields("reproductive_guild").Value = rg
    
    
    TBL.Fields("fecundity").Value = fecundity_v
    TBL.Fields("fecundity_1st").Value = fecundity_v1
    TBL.Fields("fecundity_2nd").Value = fecundity_v2
    'TBL.Fields("fecundity_text").Value = fecundity_text
    
    
    
    TBL.Fields("Emsy").Value = vemsy
    TBL.Fields("Eopt").Value = veopt
    TBL.Fields("Fmsy").Value = vfmsy
    TBL.Fields("Fopt").Value = vfopt
    TBL.Fields("Lc").Value = vLc
    TBL.Fields("Lc_type").Value = Lc_lt
    TBL.Fields("E").Value = vE
    TBL.Fields("YR").Value = vYR
    
    TBL.Fields("resilience").Value = resiliency
    
    
    
    TBL.Fields("rm").Value = Round(vrm, 2)
    TBL.Fields("Lr").Value = vlr
    TBL.Fields("Lr_type").Value = lr_lt
    
    
    TBL.Fields("main_food").Value = mf
    TBL.Fields("trophic_level").Value = tl
        
     
     
     
     
    
    
    TBL.Fields("QB").Value = finalqb
    TBL.Fields("QB_text").Value = finalqb_text
    TBL.Fields("QB_winf").Value = vwinf
    TBL.Fields("QB_temp").Value = var_temp_qb
    TBL.Fields("QB_A").Value = vAfin
    
    
    
    
    TBL.Update
        
        
    End If
    TBL.MoveNext
Wend


'MsgBox TBL.RecordCount
'MsgBox i
TBL.Close
MDB.Close



End Sub

Private Sub Command6_Click()
'2v2


Dim MDB As Database
Dim TBL As Recordset

Set MDB = OpenDatabase(App.Path & "\keyfacts.mdb")
Set TBL = MDB.OpenRecordset("SELECT Matrix.*, SPECIES.FamCode, SPECIES.Length, SPECIES.LengthFemale, SPECIES.LTypeMaxM, SPECIES.LTypeMaxF " & _
"FROM Matrix LEFT JOIN SPECIES ON Matrix.SpecCode = SPECIES.SpecCode;", dbOpenDynaset)

Dim i As Long
i = 0
Dim X As Long
Dim SSS As Variant
Dim vkf_comment As String
Dim STR As String




TBL.MoveFirst
While Not TBL.EOF

'start initialize c6
    Variable_Length = Null
    variable_type1 = Null
    variable_infinity = Null
    variable_type2 = Null
    linf_r1 = Null
    Linf_r2 = Null
    finalk = Null
    variable_q = Null
    xto = Null
    var_temp = Null
    final_mortality = Null
    m1st = Null
    m2nd = Null
    m_comment = Null
    lspan = Null
    lspan_r1 = Null
    lspan_r2 = Null
    GENTIME = Null
    gentime_r1 = Null
    gentime_r2 = Null
    gtime = Null
    gtime_r1 = Null
    gtime_r2 = Null
    vlmaturity = Null
    lm_1 = Null
    lm_2 = Null
    maturity_lt = Null
    lmaxyield = Null
    lmaxyield_range1 = Null
    lmaxyield_range2 = Null
    yield_lt = Null
    lmaxyield_est = Null
    var_a = Null
    var_b = Null
    finalval = Null
    finalval_type = Null
    variable_length2 = Null
    variable_type3 = Null
    nitrogen = Null
    protein = Null
    np_weight = Null
    rg = Null
    fecundity_v = Null
    fecundity_v1 = Null
    fecundity_v2 = Null
    vemsy = Null
    veopt = Null
    vfmsy = Null
    vfopt = Null
    vLc = Null
    Lc_lt = Null
    vE = Null
    vYR = Null
    resiliency = Null
    vrm = Null
    vlr = Null
    lr_lt = Null
    mf = Null
    tl = Null
    finalqb = Null
    finalqb_text = Null
    vwinf = Null
    var_temp_qb = Null
    vAfin = Null

'end initialize

    If TBL!kf_type = 2 Then
    'If TBL!Speccode = 2 Or TBL!Speccode = 2 Then
        i = i + 1
        X = TBL!SpecCode
        var_tmax = TBL!tm_for_KF
        
    STR = "SELECT STOCKS.Stockcode From stocks " & _
    "where stocks.SpecCode=" & X & "and stocks.level = 'species in general'"
    Set getstockcode = MDB.OpenRecordset(STR, dbOpenDynaset)
    If getstockcode.RecordCount <> 0 Then
        vstockcode = getstockcode!StockCode
    Else
        vstockcode = 0
    End If
        
        
        
    '############################################################################
    '####### start main #########################################################
    '############################################################################
    var_temp = 0
    
    linf_r1 = Null
    Linf_r2 = Null
    xto = Null
    
    vlmaturity2 = Null
    gtime = Null
    '1 - vlmaturity2 / Linf_r1) / (gtime - xto), 2)







    If TBL!ltypemaxm <> "" And Not IsNull(TBL!ltypemaxm) Then
        variable_type1 = TBL!ltypemaxm
    Else
        variable_type1 = ""
    End If

    variable_type2 = variable_type1
    variable_type3 = variable_type1
    


  
    
    
    
    
    
    
    
    
    '############################################################################
    'MsgBox "111"
    
        
If (TBL!length) = 0 Then
    Variable_Length = 4
Else
    Variable_Length = TBL!length
End If

      
    
    
    
    '############################################################################
    
    
    '############################################################################
    'start linf for KF_type = 3
    
    
        If TBL!length <> "" And Not IsNull(TBL!length) Then
            variable_lmax = TBL!length
            variable_infinity = 10 ^ (0.044 + 0.9841 * (Log(TBL!length) / Log(10)))
        Else
            If TBL!lengthfemale <> "" And Not IsNull(TBL!lengthfemale) Then
                variable_lmax = TBL!lengthfemale
                variable_infinity = 10 ^ (0.044 + 0.9841 * (Log(TBL!lengthfemale) / Log(10)))
            
            Else
                variable_lmax = 1
                variable_infinity = 10 ^ (0.044 + 0.9841 * (Log(1) / Log(10)))
            End If
        End If
   
    variable_infinity = Round(variable_infinity, 1)
    vlmaturity = Round(10 ^ (0.898 * (Log(variable_infinity) / Log(10)) - 0.0781), 1)
    mat_r1 = Round(10 ^ (0.898 * (Log(variable_infinity) / Log(10)) - 0.0781 - 0.127), 1)
    mat_r2 = Round(10 ^ (0.898 * (Log(variable_infinity) / Log(10)) - 0.0781 + 0.127), 1)

      

    'If GetQprime.RecordCount = 0 Then

    
    'End If
    
    
    
    'end linf for KF_type = 1
    '############################################################################
    'MsgBox "222"
    
    '############################################################################
    'start K for KF_type = 4
    
STR = "select estimate.k from estimate where estimate.speccode = " & TBL!SpecCode
Set getk = MDB.OpenRecordset(STR, dbOpenDynaset)


'MsgBox "222 a"
vkf_comment = ""

finalk = Null
If getk.RecordCount <> 0 Then

    vkf_comment = "Values shown below are defaults; and K is an estimate from the " & _
    "Family where the species belongs. Please double-check, " & _
    "replace with better values as appropriate, and 'Recalculate'."

    If getk!k = "" Or getk!k = Null Then
        vkf_comment = "No value for K."
    Else
        finalk = Round(getk!k, 3)
    End If
    
    

Else

    'MsgBox "no k in estimate = " & X

vkf_comment = "Values shown below are defaults. Please double-check, " & _
"replace with better values as appropriate, and 'Recalculate'."

End If

'MsgBox "222 b"


If IsNull(finalk) Then
    xto = Null
Else
    xto = Round(-1 * (10 ^ (-0.3922 - 0.2752 * Log(variable_infinity) / Log(10) - 1.038 * Log(finalk) / Log(10))), 2)
End If


'MsgBox getk.RecordCount


    

    





'//'for variable_infinity
variable_infinity = Round(variable_infinity, 1)
        X01 = Log(variable_infinity) / Log(10)
        variable_lmax = X01 - 0.044
        variable_lmax = variable_lmax / 0.9841
        variable_lmax = Round((10 ^ variable_lmax), 1)

'MsgBox "222 c"

iclarm = Round((10 ^ ((Log(variable_infinity) / Log(10) - 0.044) / 0.9841)), 1)
linf_r1 = Round((10 ^ (0.044 + 0.9841 * Log(iclarm) / Log(10) - 0.074)), 1)
Linf_r2 = Round((10 ^ (0.044 + 0.9841 * Log(iclarm) / Log(10) + 0.074)), 1)

'<input type="text" name="Linf_r1"
'value="#numberformat(evaluate(10^(0.044 + 0.9841 * LOG10(val(variable_length)) - 0.074) ),"9999.9")#"
'size="6" align="right" onFocus="noedit(1)"  >
'</font>
'-
'<font face="fixedsys">
'<input type="text" name="Linf_r2"
'value="#numberformat(evaluate(10^(0.044 + 0.9841 * LOG10(val(variable_length)) + 0.074) ),"9999.9")#"
'size="6" align="right" onFocus="noedit(1)"  ></font>




'//'for vlmaturity
vlmaturity = Round((10 ^ (0.898 * Log(variable_infinity) / Log(10) - 0.0781)), 1)
mat_r1 = Round((10 ^ (0.898 * Log(variable_infinity) / Log(10) - 0.0781 - 0.127)), 1)
mat_r2 = Round((10 ^ (0.898 * Log(variable_infinity) / Log(10) - 0.0781 + 0.127)), 1)
vlmaturity2 = vlmaturity

'MsgBox "222 d"

'//'for lmyield
b = (10 ^ 1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742)
lmaxyield = Round(b, 1)
lmaxyield_range1 = Round((10 ^ 1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742 - 0.073), 1)
lmaxyield_range2 = Round((10 ^ 1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742 + 0.073), 1)

'MsgBox "222 e"
    'end K for KF_type = 3
    '############################################################################
'MsgBox "333"
    
    
    
    
    '############################################################################
    'start var_temp
    
    If var_temp = 0 Then
        STR = "SELECT Avg(POPGROWTH.Temperature) AS AvgOfT " & _
        "From POPGROWTH " & _
        "WHERE (((POPGROWTH.SpecCode)=" & TBL!SpecCode & ") AND (Not (POPGROWTH.Temperature)=0 " & _
        "And (POPGROWTH.Temperature) Is Not Null))"
        Set avgt = MDB.OpenRecordset(STR, dbOpenDynaset)
        
        While Not avgt.EOF
            var_temp = avgt!avgoft
            avgt.MoveNext
        Wend
    End If
    
    
    
    
    If var_temp = 0 Then
    STR = "SELECT Avg(([tempmin]+[tempmax])/2) AS vAvg " & _
    "From STOCKS " & _
    "WHERE   (   ((stocks.tempmin)<>0 And (stocks.tempmin) Is Not Null) AND " & _
                "((stocks.tempmax)<>0 And (stocks.tempmax) Is Not Null) AND " & _
                "((stocks.Speccode)=" & TBL!SpecCode & ")    )"
    Set opt_stocks = MDB.OpenRecordset(STR, dbOpenDynaset)
    If opt_stocks!vAvg > 0 Then
        var_temp = opt_stocks!vAvg
    End If
    End If
    
    
'    MsgBox "444"
If IsNull(var_temp) Then
    STR = "SELECT STOCKS.EnvTemp From STOCKS " & _
    "WHERE (((STOCKS.stockcode)=" & vstockcode & "));"
    Set opt_stocks = MDB.OpenRecordset(STR, dbOpenDynaset)
    
    While Not opt_stocks.EOF
        If opt_stocks!envtemp = "boreal" Then
            var_temp = 6
        ElseIf opt_stocks!envtemp = "deep-water" Then
            var_temp = 8
        ElseIf opt_stocks!envtemp = "high altitude" Then
            var_temp = 1
        ElseIf opt_stocks!envtemp = "polar" Then
            var_temp = 1
        ElseIf opt_stocks!envtemp = "subtropical" Then
            var_temp = 17
        ElseIf opt_stocks!envtemp = "temperate" Then
            var_temp = 10
        ElseIf opt_stocks!envtemp = "tropical" Then
            var_temp = 25
        End If
        opt_stocks.MoveNext
    Wend
End If
   
    
    'end var_temp
    '############################################################################
    
    
    '############################################################################
    'start variable_type2
    
    
    STR = "SELECT POPgrowth.Speccode,popgrowth.temperature From POPgrowth " & _
    "WHERE (((POPgrowth.Speccode)=" & X & "))"
    Set withgrowth = MDB.OpenRecordset(STR, dbOpenDynaset)
    
    variable_type2 = variable_type1
    
    
    
    'end variable_type2
    '############################################################################
    
    
    '############################################################################
    'start M and others
    eli = X
    
    m1st = Null
    m2nd = Null
    m_comment = Null
        


        
        
    
    If variable_type2 = "TL" Then
    'If 1 = 1 Then
    
        If IsNull(finalk) Or finalk = "" Then
            final_mortality = Round(10 ^ (0.566 - 0.718 * (Log(variable_infinity) / Log(10)) + 0.02 * var_temp), 2)
            m1st = 10 ^ (0.566 - 0.718 * (Log(variable_infinity) / Log(10)) + 0.02 * var_temp - 0.249)
            m2nd = 10 ^ (0.566 - 0.718 * (Log(variable_infinity) / Log(10)) + 0.02 * var_temp + 0.249)
            m_comment = "Estimated from Linf. and annual mean temp."
        Else
            final_mortality = Round((10 ^ (0.333 - 0.246 * Log(variable_infinity) / Log(10) + 0.744 * Log(finalk) / Log(10) + 0.01 * var_temp)), 2)
            m1st = Round((10 ^ (0.333 - 0.246 * Log(variable_infinity) / Log(10) + 0.744 * Log(finalk) / Log(10) + 0.01 * var_temp - 0.18)), 2)
            m2nd = Round((10 ^ (0.333 - 0.246 * Log(variable_infinity) / Log(10) + 0.744 * Log(finalk) / Log(10) + 0.01 * var_temp + 0.18)), 2)
            m_comment = "Estimated from Linf., K and annual mean temp."
            
        End If
        
        
        If X = 673 Then
        
    '    MsgBox variable_infinity
    '    MsgBox var_temp
        
        End If
        
    Else
        final_mortality = Null
        'm1st = ""
        'm2nd = ""
    End If
    
    
  
    
    
    
    
    
    
    
    'end M and others
    '############################################################################
    
    
'############################################################################
    'start tm
    
    'old
    'gtime = Round(xto + (-1 * (Log(1 - vlmaturity / variable_infinity) / finalk)), 1)
    'gtime_r1 = Round(xto + (-1 * (Log(1 - mat_r1 / Linf_r1) / finalk)), 1)
    'gtime_r2 = Round(xto + (-1 * (Log(1 - mat_r2 / Linf_r2) / finalk)), 1)
    
    
'//  'for gtime

If (1 - vlmaturity / variable_infinity > 0) Then
    gtime = Round(xto + (-1 * (Log(1 - vlmaturity / variable_infinity) / finalk)), 1)
    gtime_r1 = Round(xto + (-1 * (Log(1 - mat_r1 / linf_r1) / finalk)), 1)
    gtime_r2 = Round(xto + (-1 * (Log(1 - mat_r2 / Linf_r2) / finalk)), 1)
    
    'If X = 673 Then
    '    MsgBox "xto = " & xto
    '    MsgBox "vlmaturity = " & vlmaturity
    '    MsgBox "variable_infinity = " & variable_infinity
    '    MsgBox "finalk = " & finalk
    '    MsgBox "gtime = " & gtime
    'End If
    
End If
    
    'end tm
    '############################################################################
    
    
    
    '############################################################################
    'start life span
    
    
If IsNull(finalk) Then
Else
    'lspan = (3 / finalk) + xto
    lspan = Round(((3 / finalk) + xto), 1)
End If
    
    
lspan_r1 = Null
lspan_r2 = Null
   

xk1 = Null
xk2 = Null
    
'home comment
If Not IsNull(finalk) Then
If Not IsNull(gtime) Then
    If (1 - vlmaturity2 / linf_r1) > 0 Then
        xk1 = Round(-Log(1 - vlmaturity2 / linf_r1) / (gtime - xto), 2)
        xk2 = Round(-Log(1 - vlmaturity2 / Linf_r2) / (gtime - xto), 2)
        lspan_r1 = Round(((3 / xk1) + (xto)), 1)
        lspan_r2 = Round(((3 / xk2) + (xto)), 1)
    End If
End If
End If

    'end life span
    '############################################################################
    
    
    '############################################################################
    'start gentime
    GENTIME = Null
        
varlopt = Round(10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742), 1)
evlmaturity = Round(10 ^ (0.898 * (Log(variable_infinity) / Log(10)) - 0.0781), 1)

If evlmaturity >= varlopt Then
    lm100 = evlmaturity + (variable_infinity - evlmaturity) / 4
    GENTIME = xto + (-1 * (Log(1 - lm100 / variable_infinity) / finalk))
    
    'If X = 673 Then
    '    MsgBox "111222"
    'End If
Else
       
lmaxyield_range1 = Round(10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742 - 0.073), 1)
lmaxyield_range2 = Round(10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742 + 0.073), 1)


If IsNull(finalk) Or finalk = "" Then
Else
    GENTIME = xto + (-1 * (Log(1 - varlopt / variable_infinity) / finalk))
    gentime_r1 = xto + (-1 * (Log(1 - lmaxyield_range1 / variable_infinity) / finalk))
    gentime_r2 = xto + (-1 * (Log(1 - lmaxyield_range2 / variable_infinity) / finalk))
End If
       
    
    'If X = 673 Then
    '    MsgBox "333444 = " & GENTIME
    'End If

End If
    
    'end gentime
    '############################################################################
    
    
    
    
        
    
    
    
    
    
        
        
        
        
        
        
    '############################################################################
    'start Lm se
    lm_1 = 10 ^ (0.898 * (Log(variable_infinity) / Log(10)) - 0.0781 - 0.127)
    lm_2 = 10 ^ (0.898 * (Log(variable_infinity) / Log(10)) - 0.0781 + 0.127)
    maturity_lt = variable_type2
    'end Lm se
    '############################################################################
    
        
    '############################################################################
    'start lopt
    
    If xxxLoo <> 0 Then '<!--- meaning with growth, median Qprime --->
        qtest1 = variable_infinity * (3 / (3 + final_mortality / finalk))
        qtest2 = 10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742)
        If qtest1 < qtest2 Then
            lmaxyield = 10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742)
            'MsgBox "2"
        Else
            lmaxyield = variable_infinity * (3 / (3 + Round(final_mortality, 2) / Round(finalk, 2)))
            'MsgBox "3"
            'MsgBox lmaxyield
            'MsgBox variable_infinity
            'MsgBox final_mortality
            'MsgBox finalk
        End If
    Else
        lmaxyield = 10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742)
        'MsgBox "4"
    End If
    
    
    'this is from kf_type=1
    'If xxxLoo = 0 Then '<!--- meaning w/o growth, median Qprime --->
    '    lmaxyield_est = "Estimated from Linf."
    'Else
    
        'ito ay parang pang kf_type=1
        'If qtest1 < qtest2 Then
        '    lmaxyield_range1 = 10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742 - 0.073)
        '    lmaxyield_range2 = 10 ^ (1.0421 * (Log(variable_infinity) / Log(10)) - 0.2742 + 0.073)
        '    lmaxyield_est = "Estimated from Linf."
        'Else
        '    lmaxyield_range1 = Null
        '    lmaxyield_range2 = Null
        '    lmaxyield_est = "Estimated from Linf., K and M."
        'End If
    'End If
    
    
    
    
    
    
        
    yield_lt = variable_type2
    
    'end lopt
    '############################################################################
    
    
    '############################################################################
    'start of l-w
    variable_length2 = Round(variable_infinity, 1)


    'start <!--- get the median of LW --->
    STR = "SELECT  POPLW.SpecCode,POPLW.LengthMin,POPLW.LengthMax,POPLW.Number,POPLW.Sex, POPLW.a, POPLW.b, COUNTREF.paese, " & _
        "poplw.autoctr, poplw.locality,poplw.type, poplw.a , poplw.b FROM COUNTREF INNER JOIN POPLW ON COUNTREF.C_Code = POPLW.C_Code " & _
    "WHERE (((POPLW.SpecCode)=" & TBL!SpecCode & "))        order by poplw.b"
    Set medlwb = MDB.OpenRecordset(STR, dbOpenDynaset)
    
    If medlwb.RecordCount <> 0 Then
        xLW = (medlwb.RecordCount / 2) + 0.5
        xLW = Round(xLW)
        
        'If X = 68 Then
        '    MsgBox xLW
        'End If
        
        ii = 0
        While Not medlwb.EOF
            ii = ii + 1
            If ii = xLW Then
                v_a = medlwb!a
                v_b = medlwb!b
                lwtype = medlwb!Type
            End If
            medlwb.MoveNext
        Wend
    End If
    'end <!--- get the median of LW --->



'<!--- start of lw --->
If medlwb.RecordCount <> 0 Then
    'start parang hindi dada-anan ito
    'if #parameterexists(variable_length2)# is "No">
    '    variable_length2 = #numberformat(length,"9999.9")#>
    'end if
    'end parang hindi dada-anan ito
    
    
    'if #parameterexists(var_a)# is "No">
    var_a = v_a
    'end if
    'if #parameterexists(var_b)# is "No">
    var_b = v_b
    'end if
    finalval = ((variable_length2 ^ Round(var_b, 3)) * Round(var_a, 4))
    'finalval = ((46 ^ 3.13) * 0.0054)
    
    'If X = 68 Then
    '    MsgBox variable_length2 & " " & var_b & " " & var_a & " = " & finalval
    'End If
    
    variable_type3 = lwtype
    eli01 = Len(Trim(Round(finalval, 1))) * 2


    'If TBL!speccode = 2 Then
    '    MsgBox var_a
    '    MsgBox var_b
    '    MsgBox finalval
    'End If

    
    np_weight = finalval
    
    If finalval > 200 Then
        If finalval > 20000 Then
            finalval = Round(finalval / 1000, 1)
            finalval_type = "kg"
        Else
            finalval = Round(finalval)
            finalval_type = "g"
        End If
    Else
        finalval = Round(finalval, 1)
        finalval_type = "g"
    End If
    
       
    'W = finalval
    'a = Var_a
    'b = Var_b
   
    'Nitrogen & protein:
    
    '<!--- start new weight --->
    If finalval <> 0 Then
        'NP_weight = finalval   .... transferred up to preserve the decimal places..
        
        'If Trim(finalval_type) = "kg" Then
        '    NP_weight = NP_weight * 1000
        'End If
        
        
        
        nitrogen = 10 ^ (1.03 * (Log(np_weight) / Log(10)) - 1.65)
        nitrogen = Round(nitrogen, 1)
        protein = Round(6.25 * nitrogen, 1)
        np_weight = Round(np_weight)
        
        
        
    Else
        np_weight = Null
        nitrogen = Null
        protein = Null
    End If
    '<!--- end new weight --->
    
Else
    '<!--- don't show --->
End If

'<!--- end of lw --->
    
    
    'end of l-w
    '############################################################################
    
    
    '############################################################################
    'start reproductive_guild
    rg = ""
    If withgrowth.RecordCount <> 0 Then
    
    'MsgBox withgrowth.RecordCount
    'MsgBox "riz = " & vstockcode
    
    If withgrowth.RecordCount <> 0 Then
    
    
    
    STR = "SELECT REPRODUC.RepGuild1, REPRODUC.RepGuild2 From REPRODUC " & _
    "WHERE (((REPRODUC.StockCode)=" & vstockcode & "))"
    Set getguild = MDB.OpenRecordset(STR, dbOpenDynaset)
    rg = ""
    If getguild.RecordCount <> 0 Then
        If Trim(getguild!repguild1) = "" Then
            rg = " "
        Else
            rg = getguild!repguild1
        End If
        rg = rg + ": "
        If Trim(getguild.repguild2) = "" Then
            rg = rg & " "
        Else
            rg = rg & getguild!repguild2
        End If
        rg = Trim(rg)
    End If
    End If
    End If
    'end reproductive_guild
    '############################################################################
    

    
    
    
    '############################################################################
    'start fecundity
    
    fecundity_v = Null
    fecundity_v1 = Null
    fecundity_v2 = Null
    fecundity_text = Null
    
    
    
    If withgrowth.RecordCount <> 0 And vstockcode <> 0 Then
    
    'If X = 69 Then
    'MsgBox X
    'MsgBox Vstockcode
    'End If

STR = "SELECT Min(SPAWNING.FecundityMin) AS MinOfFecundityMin From SPAWNING " & _
"GROUP BY SPAWNING.StockCode, SPAWNING.SpecCode " & _
"HAVING ((Not (Min(SPAWNING.FecundityMin))=0) AND ((SPAWNING.StockCode)=" & vstockcode & ") AND ((SPAWNING.SpecCode)=" & X & ")) " & _
"ORDER BY Min(SPAWNING.FecundityMin)"
Set getmin = MDB.OpenRecordset(STR, dbOpenDynaset)

STR = "SELECT Max(SPAWNING.FecundityMax) AS MaxOfFecundityMax From SPAWNING GROUP BY SPAWNING.StockCode, SPAWNING.SpecCode " & _
"HAVING (((SPAWNING.StockCode)=" & vstockcode & ") AND ((SPAWNING.SpecCode)=" & X & ") AND (Not (Max(SPAWNING.FecundityMax))=0))"
Set getmax = MDB.OpenRecordset(STR, dbOpenDynaset)
    
    
If getmin.RecordCount <> 0 Or getmax.RecordCount <> 0 Then
    
    y01 = 0
    y02 = 0
    fecundity_v1 = ""
    fecundity_v2 = ""
    
    
    If getmin.RecordCount <> 0 And getmax.RecordCount <> 0 Then
    If getmin!minoffecunditymin > 0 And getmax!maxoffecunditymax > 0 Then
    y01 = Log(getmin.minoffecunditymin) / Log(10)
    y02 = Log(getmax.maxoffecunditymax) / Log(10)
    y01 = y01 + y02
    y01 = y01 / 2
    fecundity_v = Round(10 ^ y01)
    End If
    End If
    
    If getmin.RecordCount <> 0 Then
        If getmin!minoffecunditymin = "" Then
            fecundity_v1 = "no value (min.)"
        Else
            fecundity_v1 = Round(getmin!minoffecunditymin)
        End If
    Else
        fecundity_v1 = "no record (min.)"
    End If
    
    
    If getmax.RecordCount <> 0 Then
        If getmax!maxoffecunditymax = "" Then
            fecundity_v2 = "no value (max.)"
        Else
            fecundity_v2 = Round(getmax!maxoffecunditymax)
        End If
    Else
        fecundity_v2 = "no record (max.)"
    End If
    
    
    
    If getmin.RecordCount <> 0 And getmax.RecordCount <> 0 Then
        fecundity_text = "Estimated as geometric mean."
    End If


End If
    End If
    'end fecundity
    '############################################################################
    
    
    
'############################################################################
'start yrecruit

'<!--- start yrecruit --->

    vemsy = Null
    veopt = Null
    vfmsy = Null
    vfmsy_rm = Null
    vfopt = Null
    vLc = Null
    Lc_lt = Null
    vE = Null
    vYR = Null





vE = 0.5
vLc = Round(0.4 * variable_infinity, 1)
If final_mortality > 0 Then
    
    'If X = 942 Then
    '    MsgBox "meron vYR"
    'End If
    
    vU = 1 - vLc / Round(variable_infinity, 2)
    MK = Round(final_mortality, 2) / Round(finalk, 2)
    
    vx1 = 0
    vx2 = 0
    firstloop = "Y"
    
    oldy = 0
    vlope = 0
    vemsy = 0
    veopt = 0
    vfmsy = 0
    vfmsy_rm = 0
    vfopt = 0
    
    
    While vx1 <= 1
        
        vx1 = vx2
        vx2 = vx2 + 0.001

        vm1 = (1 - vx1) / MK
        vy1 = vx1 * (vU ^ MK) * (1 - ((3 * vU) / (1 + vm1)) + ((3 * vU ^ 2) / (1 + 2 * vm1)) - ((vU ^ 3) / (1 + 3 * vm1)))
        vm2 = (1 - vx2) / MK
        vy2 = vx2 * (vU ^ MK) * (1 - ((3 * vU) / (1 + vm2)) + ((3 * vU ^ 2) / (1 + 2 * vm2)) - ((vU ^ 3) / (1 + 3 * vm2)))
        vslope = (vy2 - vy1) / (vx2 - vx1)
        
        If oldy <> 0 Then
            If vy1 >= oldy Then
            Else
                vemsy = Round(vx1 - 0.001, 2)
                '<cfbreak>
                GoTo EliGo
            End If
        End If
        
        
        oldy = vy1
        
        If firstloop = "Y" Then
            firstvalue = (vy2 - vy1) / (vx2 - vx1)
            firstloop = "N"
        End If

        
    
        If veopt = 0 Then
        If Round(vslope, 3) = Round(firstvalue / 10, 3) Then
            veopt = vx1
        End If
        End If
        
        
    Wend
EliGo:
    '<!--- end get e  --->
    

    vm = (1 - vE) / MK
    lijosh = vE * (vU ^ MK) * (1 - ((3 * vU) / (1 + vm)) + ((3 * vU ^ 2) / (1 + 2 * vm)) - ((vU ^ 3) / (1 + 3 * vm)))
    '<input type="text" name="vYR" value=round(lijosh,'9.9999')#" size="6" onFocus="noedit(2)"  align="right">
    vYR = Round(lijosh, 4)

Else
    '<input type="text" name="vYR" value="" size="6" onFocus="noedit(2)"  align="right">
    vYR = Null
    
    'If X = 942 Then
    '    MsgBox "vYR is null"
    'End If
    
End If

    vLc = Round(vLc, 1)
    Lc_lt = Trim(variable_type2)
    e = vE = Round(vE, 2)
    
    If final_mortality > 0 Then
    
        veopt = Round(veopt, 2)
        vemsy = Round(vemsy, 2)
        'vfmsy = Round(final_mortality * vEmsy / (1 - vEmsy), 2)
        vfmsy = final_mortality * vemsy / (1 - vemsy)
        vfmsy_rm = final_mortality * vemsy / (1 - vemsy)
        'vFopt = Round(final_mortality * vEopt / (1 - vEopt), 2)
        vfopt = final_mortality * veopt / (1 - veopt)
        
        
        
        
        
        
        veopt = Round(veopt, 2)
        vemsy = Round(vemsy, 2)
        vfmsy = Round(vfmsy, 2)
        vfopt = Round(vfopt, 2)
    Else
        vemsy = Null
        veopt = Null
        vfmsy = Null
        vfmsy_rm = Null
        vfopt = Null
    End If





'<!--- end yrecruit   width  msy --->


'end yrecruit
'############################################################################
    
    
    '############################################################################
    'start resiliency
    resiliency = Null
    
    
If finalk <= 0.05 Or var_tmax > 30 Then
    resiliency = "Very low; decline threshold 0.70"
ElseIf finalk <= 0.15 Or var_tmax >= 11 Then
    resiliency = "Low; decline threshold 0.85"
ElseIf finalk <= 0.3 Or var_tmax >= 4 Then
    resiliency = "Medium; decline threshold 0.95"
ElseIf finalk > 0.3 Or var_tmax < 4 Then
    resiliency = "High; decline threshold 0.99"
Else
    resiliency = "Please enter values for K, tmax."
End If
    
    
    
    
    
    
    
    'If Round(finalk, 2) <= 0.05 Then
    '    resiliency = "Very low; decline threshold 0.70"
    'ElseIf Round(finalk, 2) <= 0.15 Then
    '    resiliency = "Low; decline threshold 0.85"
    'ElseIf Round(finalk, 2) <= 0.3 Then
    '    resiliency = "Medium; decline threshold 0.95"
    'ElseIf Round(finalk, 2) > 0.3 Then
    '    resiliency = "High; decline threshold 0.99"
    'Else
    '    resiliency = "Please enter values for K."
    'End If
    
    
    'end resiliency
    '############################################################################
    
    
    
    
    
    
    '############################################################################
    'start rm
    vlr = Round(0.4 * variable_infinity, 1)
    vfmsy = Round(vfmsy, 2)
    vrm = Round(2 * vfmsy, 2)
    
    If vrm <> 0 Then
    vypdt = Log(2) / vrm
    End If
    
    lr_lt = Trim(variable_type2)
    
    'If X = 2 Then
    '    MsgBox vfmsy
    '    MsgBox vrm
    'End If
    
    'end rm
    '############################################################################
    
    
    
    '############################################################################
    'start eco
    STR = "SELECT  ECOLOGY.StockCode,ECOLOGY.DietTroph,ECOLOGY.DietSeTroph,ECOLOGY.FoodTroph, " & _
    "ECOLOGY.FoodSeTroph,0 as EcoTroph,0 as EcoSeTroph,'' as mainfood,ECOLOGY.herbivory2 " & _
    "From ECOLOGY WHERE (((ECOLOGY.StockCode)=" & vstockcode & "))"
    Set witheco = MDB.OpenRecordset(STR, dbOpenDynaset)
    
mf = Null
tl = Null
If witheco.RecordCount <> 0 Then
    If witheco!mainfood <> "" Then
        mf = witheco!mainfood
    End If
    If witheco!herbivory2 <> "" Then
        mf = mf & " " & witheco!herbivory2
    End If
    
    
    'Trophic level:
    tl = Null
    If witheco!dietTroph <> "" Then
        tl = Round(witheco!dietTroph, 1)
        If witheco!DietSeTroph <> 0 Then
            tl = tl & "&nbsp;&nbsp;&nbsp;"
            tl = tl & "+/- s.e. " & Round(witheco!DietSeTroph, 2)
        End If
        tl = tl & " Estimated from diet data."
    Else
        If witheco!foodTroph <> "" Then
            tl = tl & " " & Round(witheco!foodTroph, 1)
            If witheco!FoodSeTroph <> 0 Then
                tl = tl & "+/- s.e. " & Round(witheco!FoodSeTroph, 2)
            End If
            tl = tl & " Estimated from food data."
        Else
            If witheco!EcoTroph <> "" Then
                tl = tl & " " & Round(witheco!EcoTroph, 1)
                If witheco!EcoSeTroph <> 0 Then
                    tl = tl & "+/- s.e. " & Round(witheco!EcoSeTroph, 2)
                End If
                tl = tl & " Estimated from Ecopath model."
            End If
        End If
    End If
End If


'<!---start of food consumption  </cfif --->

    
    'end eco
    '############################################################################
    
    
    
    
    
    
    
    
    '############################################################################
    'start qb
    
    finalqb = Null
    finalqb_text = Null
    vwinf = Null
    var_temp_qb = Null
    vAfin = Null
    
    
    
    
    
'<!--- start food consumption --->
contqb = "Y"
'<!--- start get median popqb --->
STR = "SELECT popqb.popqb FROM popqb WHERE (((popqb.speccode)=" & X & ")) ORDER BY popqb.popqb"
Set qbmedian = MDB.OpenRecordset(STR, dbOpenDynaset)


'start of old
'vmedian = (qbmedian.RecordCount / 2) + 0.5
'vmedian = Round(vmedian)
'end of old


vMedian = (qbmedian.RecordCount / 2) + 0.5
int_part = Int(vMedian)
str_conv = "" & vMedian

If Len(str_conv) = 1 Then
    vMedian = Round(vMedian)
Else

    If Mid(str_conv, Len(str_conv) - 1, 2) = ".5" Then
        vMedian = int_part + 1
    Else
        vMedian = Round(vMedian)
    End If
End If




vpopqb = 0


'If X = 69 Then
'    MsgBox qbmedian.RecordCount
'    MsgBox (qbmedian.RecordCount / 2) + 0.5
'    MsgBox Int((qbmedian.RecordCount / 2) + 0.5)
'    MsgBox Round(vmedian)
'    MsgBox Round(vmedian, 1)
'End If


ii = 0
While Not qbmedian.EOF

    ii = ii + 1
    If ii = vMedian Then
        vpopqb = qbmedian!popqb
    End If
    qbmedian.MoveNext
Wend


'<!--- end get median popqb --->

If vpopqb > 0 Then
    explain = "with popqb record"
    contqb = "N"
Else
    explain = "no popqb record"
    
    '<!---#############################################################################################--->
    '<!---### start of no popqb #######################################################################--->
    '<!---#############################################################################################--->
    If medlwb.RecordCount# > 0 Then
        explain = explain & "; with lw rel"
    Else
        explain = explain & "; no lw rel"
    End If

    '<!--- start of A --->
    STR = "SELECT Swimming.AspectRatio as aspectratio From Swimming WHERE (((Swimming.SpecCode)=" & X & "));"
    Set getAR = MDB.OpenRecordset(STR, dbOpenDynaset)
    vAfin = 0
    If getAR.RecordCount > 0 Then
       While Not getAR.EOF
            If getAR!aspectratio > 0 Then
                vAfin = getAR!aspectratio
                explain = explain & "; with aspect ratio"
            Else
                vAfin = 0
                explain = explain & "; w/o aspect ratio"
            End If
            getAR.MoveNext
        Wend
    Else
        explain = explain & "; w/o aspect ratio"
    End If
    '<!--- end of A --->
    
    
    '<!--- start of h and d --->
    SelectHD = "N"
    '<!---
    '    go ecology feeding type
    '    if none troph diettroph
    '        if none troph foodtroph items
    '            if none troph ecotroph
    '--->


    STR = "SELECT ECOLOGY.Herbivory2, 0 as EcoTroph, ECOLOGY.DietTroph, ECOLOGY.FoodTroph From ECOLOGY " & _
    "WHERE (((ECOLOGY.StockCode)=" & vstockcode & "))"
    Set getft = MDB.OpenRecordset(STR, dbOpenDynaset)

    
    
    If getft.RecordCount <> 0 Then
    If Trim(getft!herbivory2) <> "" Then
    
        If Trim(getft!herbivory2) = "mainly animals (troph. 2.8 and up)" Then
            vh = 0
            vd = 0
        ElseIf Trim(getft!herbivory2) = "mainly animals (troph. 2.8 up)" Then
            vh = 0
            vd = 0
        ElseIf Trim(getft!herbivory2) = "mainly plants/detritus (troph. 2-2.19)" Then
            vh = 0
            vd = 1
        ElseIf Trim(getft.herbivory2) = "plants/detritus+animals (troph. 2.2-2.79)" Then
            vh = 1
            vd = 0
        End If
        explain = explain & "; w/ feeding type"
        WithFT = 1
    
    Else
    
    
        explain = explain & "; w/o feeding type"
        WithFT = 0
        If (getft!dietTroph) > 0 Then
            If getft!dietTroph >= 2 And getft!dietTroph <= 2.19 Then
                vh = 0
                vd = 1
            ElseIf getft!dietTroph >= 2.2 And getft!dietTroph <= 2.79 Then
                vh = 1
                vd = 0
            ElseIf getft!dietTroph >= 2.8 Then
                vh = 0
                vd = 0
            End If
            explain = explain & "; from diettroph"
        Else
            If getft!foodTroph > 0 Then
                If getft!foodTroph >= 2 And getft.foodTroph <= 2.19 Then
                    vh = 0
                    vd = 1
                ElseIf getft.foodTroph >= 2.2 And getft.foodTroph <= 2.79 Then
                    vh = 1
                    vd = 0
                ElseIf getft!foodTroph >= 2.8 Then
                    vh = 0
                    vd = 0
                End If
                explain = explain & "; from foodtroph"
            Else
                If getft!EcoTroph > 0 Then
                    If getft!EcoTroph >= 2 And getft!EcoTroph <= 2.19 Then
                        vh = 0
                        vd = 1
                    ElseIf getft!EcoTroph >= 2.2 And getft!EcoTroph <= 2.79 Then
                        vh = 1
                        vd = 0
                    ElseIf getft!EcoTroph >= 2.8 Then
                        vh = 0
                        vd = 0
                    End If
                    explain = explain & "; from ecotroph"
                Else
                    explain = explain & "; no diet,food,eco trophs; select h,d"
                    '<!--- blank means yes
                    'contqb = "Y">
                    '--->
                    SelectHD = "Y"
                    '<!---
                    'contqb = "N">
                    'cont. with the search
                    'now the genus of median of diet,food,eco
                    '--->
                End If
            End If
        End If
    End If
    End If
    '<!--- end of h and d --->

'<!---#############################################################################################--->
'<!---### end of no popqb #########################################################################--->
'<!---#############################################################################################--->
End If
'<!--- end food consumption --->









'<!---start of food consumption  </cfif --->

'<td align="left">Food consumption (Q/B):</td>
If vpopqb > 0 Then
    finalqb = Round(vpopqb, 2)
    finalqb_text = "times the body weight per year"
Else
    '<!---#############################################################################################--->
    '<!---#############################################################################################--->
    '<!---#############################################################################################--->
    If contqb = "Y" Then
        '<!--- start for winf --->
        If medlwb.RecordCount > 0 Then
            'orig
            'vwinf = finalval
            
            'use this instead as the finalval is sometimes in kg
            vwinf = np_weight
            'If X = 68 Then
            '    MsgBox "111"
            'End If
            
            whereWinf = "lw"
        Else
            vwinf = 0.01 * variable_infinity ^ 3
            'If X = 68 Then
            '    MsgBox "222"
            'End If
            
            whereWinf = "DP"
        End If
        '<!--- end for winf --->
        If SelectHD = "Y" Then
            vh = 3
            vd = 3
        End If
        'start vb comment
        '<input type="hidden" name="vh" value="#vh#" size="1" onFocus="noedit(2)"  >
        '<input type="hidden" name="vd" value="#vd#" size="1" onFocus="noedit(2)"  >
        'end vb comment
        
        If vAfin > 0 Then
            xyz = vAfin
        Else
            xyz = 1.32
        End If
        
        elix = 0
        If vh = 3 Then
            '<!---w/o h, d --->
            If vwinf <> 0 Then
                elix = 10 ^ (7.964 - 0.204 * Log(vwinf) / Log(10) - 1.965 * (1000 / (var_temp + 273.15)) + 0.083 * xyz + 0.532 * 1 + 0.398 * 0)
                elix2 = 10 ^ (7.964 - 0.204 * Log(vwinf) / Log(10) - 1.965 * (1000 / (var_temp + 273.15)) + 0.083 * xyz + 0.532 * 0 + 0.398 * 0)
                elix = (elix2 + elix) / 2
            End If
        Else
            '<!---with h, d--->
            If Not IsNull(vwinf) And vwinf <> 0 Then
            elix = 10 ^ (7.964 - 0.204 * Log(vwinf) / Log(10) - 1.965 * (1000 / (var_temp + 273.15)) + 0.083 * xyz + 0.532 * vh + 0.398 * vd)
            End If
        End If
        
        
        If whereWinf = "none" Or elix = 0 Then
            finalqb = Null
        Else
            finalqb = Round(elix, 1)
        End If
        finalqb_text = "times the body weight per year"
        
        
        'Enter Winf, temperature, aspect ratio (A), and food type to estimate Q/B
        If whereWinf = "none" Then
            'Winf =
            vwinf = Null
            'If X = 68 Then
            '    MsgBox "333"
            'End If
            
            'Temp. =
            var_temp_qb = Round(var_temp, 1)
        Else
            'Winf =
            vwinf = Round(vwinf, 1)
            
            'If X = 68 Then
            '    MsgBox "444"
            'End If
            
            'Temp. =
            var_temp_qb = Round(var_temp, 1)
        End If
        
        If vAfin > 0 Then
            'A =
            vAfin = Round(vAfin, 2)
        Else
            vAfin = 1.32
            'A =
            vAfin = Round(vAfin, 2)
            
            '<!---eeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeee--->
            'start vb comment
            '<td colspan="8"><img src="../jpgs/Tails.gif" height=29 border=0 alt=""></td>
            '<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
            '<td><input type="radio" name="eli"          onClick="finc('#contqb#',6.55)"></td>
            '<td><input type="radio" name="eli"          onClick="finc('#contqb#',1.9)"></td>
            '<td><input type="radio"  checked name="eli" onClick="finc('#contqb#',1.32)"></td>
            '<td><input type="radio" name="eli"          onClick="finc('#contqb#',1.32)"></td>
            '<td><input type="radio" name="eli"          onClick="finc('#contqb#',1.63)"></td>
            'end vb comment
            
        End If
        If SelectHD = "Y" Then
            'start vb comment
            '<input type="hidden" name="omni" value="1">
            '<td align="center"><font size="2">Detrivore</font></td>
            '<td align="center"><font size="2">Herbivore</font></td>
            '<td align="center"><font size="2">Omnivore</font></td>
            '<td align="center"><font size="2">Carnivore</font></td>
            '<td align="center"><input type="radio" name="riz"           onClick="changeHD('#contqb#',0,1)"></td>
            '<td align="center"><input type="radio" name="riz"          onClick="changeHD('#contqb#',1,0)"></td>
            '<td align="center"><input type="radio" name="riz" checked  onClick="changeHD('#contqb#',3,3)"></td>
            '<td align="center"><input type="radio" name="riz"           onClick="changeHD('#contqb#',0,0)"></td>
            '</tr></table>
            'end vb comment
            
        Else
            'start vb comment
            '<input type="hidden" name="omni" value="0">
            '<td align="center"><font size="2">Detrivore</font></td>
            '<td align="center"><font size="2">Herbivore</font></td>
            '<td align="center"><font size="2">Omnivore</font></td>
            '<td align="center"><font size="2">Carnivore</font></td>
            '<td align="center"><input type="radio" name="riz" if #vh# is 0 and #vd# is 1>checkedend if onClick="changeHD('#contqb#',0,1)"></td>
            '<td align="center"><input type="radio" name="riz" if #vh# is 1 and #vd# is 0>checkedend if onClick="changeHD('#contqb#',1,0)"></td>
            '<td align="center"><input type="radio" name="riz"                                              onClick="changeHD('#contqb#',3,3)"></td>
            '<td align="center"><input type="radio" name="riz" if #vh# is 0 and #vd# is 0>checkedend if onClick="changeHD('#contqb#',0,0)"></td>
            'end vb comment
        End If
    End If '<!--- if #contqb# is "Y"> --->
    '<!---#############################################################################################--->
    '<!---#############################################################################################--->
    '<!---#############################################################################################--->
End If
'<!--- end of food consumption  </cfif estimate width --->


   
    
    'end qb
    '############################################################################
    
    
    
    
    
    
    
    
    
    
    
        
    '############################################################################
    '####### end main ###########################################################
    '############################################################################
                
        
    TBL.Edit
    TBL.Fields("lmax").Value = Variable_Length
    TBL.Fields("lmax_type").Value = variable_type1
    
    TBL.Fields("linf").Value = variable_infinity
    TBL.Fields("linf_type").Value = variable_type2
    
    TBL.Fields("linf_1st").Value = linf_r1
    TBL.Fields("linf_2nd").Value = Linf_r2
    TBL.Fields("linf_comment").Value = "Estimated from max. length."
    
    TBL.Fields("K").Value = finalk
    TBL.Fields("PhiPrime").Value = variable_q
    TBL.Fields("to").Value = xto
    
    TBL.Fields("mean_temp").Value = var_temp
    
    TBL.Fields("M").Value = final_mortality
    TBL.Fields("M_se_1st").Value = m1st
    TBL.Fields("M_se_2nd").Value = m2nd
    TBL.Fields("m_comment").Value = m_comment
    
    TBL.Fields("life_span").Value = lspan
    'not needed here
    TBL.Fields("life_span_1st").Value = lspan_r1
    TBL.Fields("life_span_2nd").Value = lspan_r2
    
    TBL.Fields("generation_time").Value = GENTIME
    TBL.Fields("gen_time_1st").Value = gentime_r1
    TBL.Fields("gen_time_2nd").Value = gentime_r2
    
    TBL.Fields("tm").Value = gtime
    TBL.Fields("tm_1st").Value = gtime_r1
    TBL.Fields("tm_2nd").Value = gtime_r2
    
    TBL.Fields("Lm").Value = vlmaturity
    TBL.Fields("Lm_se_1st").Value = lm_1
    TBL.Fields("Lm_se_2nd").Value = lm_2
    TBL.Fields("Lm_type").Value = maturity_lt
    
    TBL.Fields("Lopt").Value = lmaxyield
    TBL.Fields("Lopt_se_1st").Value = lmaxyield_range1
    TBL.Fields("Lopt_se_2nd").Value = lmaxyield_range2
    TBL.Fields("Lopt_type").Value = yield_lt
    TBL.Fields("Lopt_text").Value = lmaxyield_est
        
    TBL.Fields("a").Value = var_a
    TBL.Fields("b").Value = var_b
    TBL.Fields("W").Value = finalval
    TBL.Fields("W_type").Value = finalval_type
    TBL.Fields("LW_length").Value = variable_length2
    TBL.Fields("LW_length_type").Value = variable_type3
    
    
    TBL.Fields("nitrogen").Value = nitrogen
    TBL.Fields("protein").Value = protein
    TBL.Fields("NitrogenProtein_weight").Value = np_weight
    
    TBL.Fields("reproductive_guild").Value = rg
    
    
    TBL.Fields("fecundity").Value = fecundity_v
    TBL.Fields("fecundity_1st").Value = fecundity_v1
    TBL.Fields("fecundity_2nd").Value = fecundity_v2
    'TBL.Fields("fecundity_text").Value = fecundity_text
    
    
    
    TBL.Fields("Emsy").Value = vemsy
    TBL.Fields("Eopt").Value = veopt
    TBL.Fields("Fmsy").Value = vfmsy
    TBL.Fields("Fopt").Value = vfopt
    TBL.Fields("Lc").Value = vLc
    TBL.Fields("Lc_type").Value = Lc_lt
    TBL.Fields("E").Value = vE
    TBL.Fields("YR").Value = vYR
    
    TBL.Fields("resilience").Value = resiliency
    
    
    
    TBL.Fields("rm").Value = Round(vrm, 2)
    TBL.Fields("Lr").Value = vlr
    TBL.Fields("Lr_type").Value = lr_lt
    
    
    TBL.Fields("main_food").Value = mf
    TBL.Fields("trophic_level").Value = tl
        
     
     
     
     
    
    
    TBL.Fields("QB").Value = finalqb
    TBL.Fields("QB_text").Value = finalqb_text
    TBL.Fields("QB_winf").Value = vwinf
    TBL.Fields("QB_temp").Value = var_temp_qb
    TBL.Fields("QB_A").Value = vAfin
    
    
    
    
    TBL.Update
        
        
    End If
    TBL.MoveNext
Wend


'MsgBox TBL.RecordCount
'MsgBox i
TBL.Close
MDB.Close



End Sub

Private Sub Command7_Click()

If Check1 Then
    Call Command1_Click
End If
If Check2 Then
    Call Command2_Click
End If
If Check3 Then
    Call Command3_Click
End If
If Check4 Then
    Call Command4_Click
End If
If Check5 Then
    Call Command5_Click
End If
If Check6 Then
    Call Command6_Click
End If
If Check7 Then
    Call Command8_Click
End If


MsgBox "Done!"


End Sub

Private Sub Command8_Click()
Dim MDB As Database
Dim TBL As Recordset

Set MDB = OpenDatabase(App.Path & "\keyfacts.mdb")
Set TBL = MDB.OpenRecordset("matrix", dbOpenDynaset)


Dim i As Long
i = 0
Dim X As Long
Dim SSS As Variant

'start var
Dim STR As String

'end var


TBL.MoveFirst
While Not TBL.EOF
    
'start initialize   c8
    vstockcode = Null
    with_growth = Null
    with_max_age_size = Null
    with_lw = Null
    with_reproduction = Null
    with_maturity = Null
    with_diet = Null
    with_food = Null
    with_food_consumption = Null
    with_spawning = Null
'end initialize
    
    
    
    X = TBL!SpecCode
    'If X = 172 Then
    
    '############################################################################
    '####### start main #########################################################
    '############################################################################
    


STR = "SELECT STOCKS.Stockcode From stocks where stocks.SpecCode=" & X & " and " & _
"(   stocks.level = 'species in general' or stocks.level = 'subspecies in general')"
Set getstockcode = MDB.OpenRecordset(STR, dbOpenDynaset)

If getstockcode.RecordCount = 0 Then
    'MsgBox X & " no stockcode"
    vstockcode = 0
Else
    vstockcode = getstockcode!StockCode
End If



with_growth = "n"
with_max_age_size = "n"
with_lw = "n"
with_reproduction = "n"
with_maturity = "n"
with_diet = "n"
with_food = "n"
with_food_consumption = "n"
with_spawning = "n"

'############################################################################

STR = "SELECT POPgrowth.Speccode,popgrowth.temperature From POPgrowth " & _
"WHERE (((POPgrowth.Speccode)=" & X & "))"
Set withgrowth = MDB.OpenRecordset(STR, dbOpenDynaset)
If withgrowth.RecordCount <> 0 Then
    with_growth = "y"
End If

'############################################################################

STR = "SELECT POPCHAR.Speccode From POPCHAR " & _
"WHERE (((POPCHAR.Speccode)=" & X & "))"
Set WithMaxSize = MDB.OpenRecordset(STR, dbOpenDynaset)
If WithMaxSize.RecordCount <> 0 Then
    with_max_age_size = "y"
End If

'############################################################################





STR = "SELECT  POPLW.SpecCode From POPLW " & _
"WHERE (((POPLW.SpecCode)=" & X & "))"
Set medlwb = MDB.OpenRecordset(STR, dbOpenDynaset)
If medlwb.RecordCount <> 0 Then
    with_lw = "y"
End If
    
'############################################################################


STR = "select reproduc.stockcode,REPRODUC.RepGuild1, REPRODUC.RepGuild2 from reproduc " & _
"where reproduc.stockcode =" & vstockcode
Set WithRepro = MDB.OpenRecordset(STR, dbOpenDynaset)



If withgrowth.RecordCount <> 0 Then
    If WithRepro.RecordCount > 0 Then
        If Trim(WithRepro!repguild1) <> "" And Trim(WithRepro.repguild2) <> "" Then
            with_reproduction = "y"
        End If
    End If
End If




'<cfif #withgrowth.recordcount# is not 0>
'    <cfif #trim(getguild.repguild1)# is not "" and #trim(getguild.repguild2)# is not "">
'        <cfif #withrepro.recordcount# GT 0>
'        </cfif>
'    </cfif>
'</cfif>









'############################################################################


STR = "SELECT MATURITY.speccode FROM MATURITY LEFT JOIN COUNTREF ON MATURITY.C_Code = COUNTREF.C_Code " & _
"WHERE (((MATURITY.Speccode)=" & X & "))"
Set WithMat = MDB.OpenRecordset(STR, dbOpenDynaset)
If WithMat.RecordCount <> 0 Then
    with_maturity = "y"
End If

'############################################################################




STR = "SELECT  ECOLOGY.StockCode,ECOLOGY.DietTroph,ECOLOGY.FoodTroph " & _
"From ECOLOGY WHERE (((ECOLOGY.StockCode)=" & vstockcode & "))"
Set witheco = MDB.OpenRecordset(STR, dbOpenDynaset)
If witheco.RecordCount <> 0 Then
If witheco!dietTroph <> "" Then
    with_diet = "y"
Else
    If witheco!foodTroph <> "" Then
        with_food = "y"
    End If
End If
End If
'############################################################################


STR = "SELECT popqb.popqb FROM popqb WHERE (((popqb.speccode)=" & X & ")) " & _
"ORDER BY popqb.popqb"
Set qbmedian = MDB.OpenRecordset(STR, dbOpenDynaset)
vMedian = (qbmedian.RecordCount / 2) + 0.5
    ii = 0
    While Not qbmedian.EOF
        ii = ii + 1
        If ii = Round(vMedian) Then
            vpopqb = qbmedian!popqb
            'can put a goto to get out of the loop
        End If
        qbmedian.MoveNext
    Wend
If vpopqb > 0 Then
    with_food_consumption = "y"
End If



'############################################################################

'start orig from keyfacts
'STR = "SELECT Min(SPAWNING.FecundityMin) AS MinOfFecundityMin From SPAWNING GROUP BY SPAWNING.StockCode, SPAWNING.SpecCode HAVING ((Not (Min(SPAWNING.FecundityMin))=0) AND ((SPAWNING.StockCode)=" & vstockcode & ") AND ((SPAWNING.SpecCode)=" & X & ")) ORDER BY Min(SPAWNING.FecundityMin)"
'Set getmin = MDB.OpenRecordset(STR, dbOpenDynaset)
'STR = "SELECT Max(SPAWNING.FecundityMax) AS MaxOfFecundityMax From SPAWNING GROUP BY SPAWNING.StockCode, SPAWNING.SpecCode HAVING (((SPAWNING.StockCode)=" & vstockcode & ") AND ((SPAWNING.SpecCode)=" & X & ") AND (Not (Max(SPAWNING.FecundityMax))=0));"
'Set getmax = MDB.OpenRecordset(STR, dbOpenDynaset)
'If getmin.RecordCount <> 0 Or getmax.RecordCount <> 0 Then
'    with_spawning = "y"
'End If
'end orig from keyfacts

STR = "SELECT spawning.speccode From SPAWNING where SPAWNING.SpecCode=" & X
Set withspawn = MDB.OpenRecordset(STR, dbOpenDynaset)

If withspawn.RecordCount <> 0 Then
    with_spawning = "y"
End If




'############################################################################


    TBL.Edit
    TBL.Fields("stockcode").Value = vstockcode
    TBL.Fields("with_growth").Value = with_growth
    TBL.Fields("with_max_age_size").Value = with_max_age_size
    TBL.Fields("with_lw").Value = with_lw
    TBL.Fields("with_reproduction").Value = with_reproduction
    TBL.Fields("with_maturity").Value = with_maturity
    TBL.Fields("with_diet").Value = with_diet
    TBL.Fields("with_food").Value = with_food
    TBL.Fields("with_food_consumption").Value = with_food_consumption
    TBL.Fields("with_spawning").Value = with_spawning
    TBL.Update
   
    
    '############################################################################
    '####### end main ###########################################################
    '############################################################################
       

    'End If
    
    
    
    TBL.MoveNext
Wend





TBL.Close
MDB.Close


End Sub

Private Sub Form_Load()
Dim STR As String




'MsgBox 10 ^ (0.566 - 0.718 * (Log(118) / Log(10)) + 0.02 * 8)

'MsgBox Round(-0.83 + (-1 * (Log(1 - 53.5 / 102.8) / 0.13)), 1)

'MsgBox 10 ^ (0.044 + 0.9841 * (Log(60 / Log(10))))

'MsgBox 10 ^ (0.044 + 0.9841 * (Log(60) / Log(10)))


'X = 6.5
'MsgBox Int(X)

'xstr = "" & X
'MsgBox Mid(xstr, Len(xstr) - 1, 2)


'MsgBox ((46 ^ 3.13) * 0.0054)



End Sub
