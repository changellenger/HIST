Attribute VB_Name = "Module1"
Public strRootDir As String
Public strLastDiaBox As String
Sub OpenProjects()
On Error Resume Next
    Hist_LOGO.Show
    strRootDir = ThisWorkBook.Path & "\" & "example" & "\"
    StrDir = ThisWorkBook.Path & "\" & "module" & "\" & "xlam" & "\"
    Workbooks("Basic.xlam").Close False
    Workbooks.Open Filename:=StrDir & "Basic.xlam"          '�������
    Workbooks("Cor.xlam").Close False
    Workbooks.Open Filename:=StrDir & "Cor.xlam"            '����м�
    Workbooks("Dm.xlam").Close False
    Workbooks.Open Filename:=StrDir & "Dm.xlam"             '�����͸��̴�
    Workbooks("Stat.xlam").Close False
    Workbooks.Open Filename:=StrDir & "Stat.xlam"           '��������
    Workbooks("Gene.xlam").Close False
    Workbooks.Open Filename:=StrDir & "StatGene.xlam"       '�������� �����ϱ�
    Workbooks("Grap.xlam").Close False
    Workbooks.Open Filename:=StrDir & "Grap.xlam"           '�׷���
    Workbooks("Qua.xlam").Close False
    Workbooks.Open Filename:=StrDir & "Qua.xlam"            'ǰ��
    Workbooks("QuaGene.xlam").Close False
    Workbooks.Open Filename:=StrDir & "QuaGene.xlam"        'ǰ�� �����ϱ�
    Workbooks("Var.xlam").Close False
    Workbooks.Open Filename:=StrDir & "Var.xlam"            '�л�м�
    Workbooks("Reg.xlam").Close False
    Workbooks.Open Filename:=StrDir & "Reg.xlam"            'ȸ�ͺм�
    Workbooks("Edu2.xlam").Close False
    Workbooks.Open Filename:=StrDir & "RegGene.xlam"        'ȸ�͵����ϱ�
    Workbooks("StatEdu.xlam").Close False
    Workbooks.Open Filename:=StrDir & "StatEdu.xlam"        '������ǥ(��.����н�)
    Workbooks("Exp2.xlam").Close False
    Workbooks.Open Filename:=StrDir & "Exp2.xlam"           '�����ȹ��
    Workbooks("Anova.xlam").Close False
    Workbooks.Open Filename:=StrDir & "Anova.xlam"          '�л�м���



End Sub




'==================Basic=====================
'������
Sub basic01(control As IRibbonControl)
    strLastDiaBox = "'Basic.xlam'!ShowfrmDisc"         '������
    Application.Run strLastDiaBox
End Sub
Sub basic02(control As IRibbonControl)
    strLastDiaBox = "'Basic.xlam'!ShowfrmFrequency" '�󵵺м�
    Application.Run strLastDiaBox
End Sub
Sub basic03(control As IRibbonControl)
    strLastDiaBox = "'Basic.xlam'!ShowframeNor" '���Լ�����
    Application.Run strLastDiaBox
End Sub
Sub basic04(control As IRibbonControl)
    strLastDiaBox = "'StatGene.xlam'!ShowframeLE_" '��л����
    Application.Run strLastDiaBox
End Sub
Sub basic05(control As IRibbonControl)
    strLastDiaBox = "'Basic.xlam'!ShowframeCrfre" '�����м�
    Application.Run strLastDiaBox
End Sub

'==================Graph=====================
Sub graph01(control As IRibbonControl)      '������
    strLastDiaBox = "'Grap.xlam'!Showscatter"
    Application.Run strLastDiaBox
End Sub
Sub graph02(control As IRibbonControl)      '������׷�
    strLastDiaBox = "'Grap.xlam'!Showhistogram"
    Application.Run strLastDiaBox
End Sub
Sub graph03(control As IRibbonControl)      '����׷���
    strLastDiaBox = "'Grap.xlam'!Showbarchart"
    Application.Run strLastDiaBox
End Sub
Sub graph04(control As IRibbonControl)      '���׷���
    strLastDiaBox = "'Grap.xlam'!ShowLinechart"
    Application.Run strLastDiaBox
End Sub
Sub graph05(control As IRibbonControl)      '���׷���
    strLastDiaBox = "'Grap.xlam'!ShowCirclechart"
    Application.Run strLastDiaBox
End Sub
Sub graph06(control As IRibbonControl)      '�����
    strLastDiaBox = "'Grap.xlam'!ShowContourline"
    Application.Run strLastDiaBox
End Sub
Sub graph07(control As IRibbonControl)      '�����׸�
    strLastDiaBox = "'Grap.xlam'!ShowInterval"
    Application.Run strLastDiaBox
End Sub
Sub graph08(control As IRibbonControl)      '���ڱ׸�
    strLastDiaBox = "'Grap.xlam'!ShowBoxchart"
    Application.Run strLastDiaBox
End Sub
Sub graph09(control As IRibbonControl)      '�ķ���׸�
    strLastDiaBox = "'Grap.xlam'!ShowParretochart"
    Application.Run strLastDiaBox
End Sub
'==================Stat=====================
Sub stat01(control As IRibbonControl)
    strLastDiaBox = "'Stat.xlam'!ShowframeOneZtest"
    Application.Run strLastDiaBox
End Sub
Sub stat02(control As IRibbonControl)
    strLastDiaBox = "'Stat.xlam'!ShowfrmOneT"
    Application.Run strLastDiaBox
End Sub
Sub stat03(control As IRibbonControl)
    strLastDiaBox = "'Stat.xlam'!ShowfrmTwoT"
    Application.Run strLastDiaBox
End Sub
Sub stat04(control As IRibbonControl)
    strLastDiaBox = "'Stat.xlam'!ShowfrmpairT"
    Application.Run strLastDiaBox
End Sub
'==================Var=====================     �л�м�
Sub var01(control As IRibbonControl)        '�Ͽ���ġ
    strLastDiaBox = "'Anova.xlam'!ShowframeFrm_1"
    Application.Run strLastDiaBox
End Sub
Sub var02(control As IRibbonControl)        '�̿���ġ
    strLastDiaBox = "'Anova.xlam'!ShowframeFrm_2"
    Application.Run strLastDiaBox
End Sub
'==================StatEdu=====================     ����н�
Sub edu00(control As IRibbonControl)        '
    strLastDiaBox = "'StatEdu.xlam'!ShowframeStNor" 'ǥ�����Ժ���
    Application.Run strLastDiaBox
End Sub
Sub edu01(control As IRibbonControl)        '
    strLastDiaBox = "'StatEdu.xlam'!ShowframeT"     'T����ǥ
    Application.Run strLastDiaBox
End Sub
Sub edu02(control As IRibbonControl)        '
    strLastDiaBox = "'StatEdu.xlam'!ShowframeF"     'F����ǥ
    Application.Run strLastDiaBox
End Sub
Sub edu03(control As IRibbonControl)        '
    strLastDiaBox = "'StatEdu.xlam'!ShowframeChi"   'ī����������ǥ
    Application.Run strLastDiaBox
End Sub
'==================Reg=====================
Sub reg01(control As IRibbonControl)
    strLastDiaBox = "'Reg.xlam'!Showframere"
    Application.Run strLastDiaBox
End Sub
Sub reg02(control As IRibbonControl)
    strLastDiaBox = "'Grap.xlam'!ShowframeReGra"
    Application.Run strLastDiaBox
End Sub
Sub reg03(control As IRibbonControl)
    strLastDiaBox = "'Reg.xlam'!Showframeglog"
    Application.Run strLastDiaBox
End Sub
Sub reg04(control As IRibbonControl)    '������� ǥ�鵵
strLastDiaBox = "'Grap.xlam'!ShowContourline"
    Application.Run strLastDiaBox
End Sub
Sub reg05(control As IRibbonControl)
    strLastDiaBox = "'Reg.xlam'!Showframeregsur"
    Application.Run strLastDiaBox
End Sub
Sub reg06(control As IRibbonControl)
    strLastDiaBox = "'Reg.xlam'!ShowStack"
    Application.Run strLastDiaBox
End Sub
'==================Cor=====================
Sub cor01(control As IRibbonControl)
    strLastDiaBox = "'Cor.xlam'!ShowframeCor"
    Application.Run strLastDiaBox
End Sub
'==================Exp=====================
Sub exp01(control As IRibbonControl)
    strLastDiaBox = "'Exp2.xlam'!showdoe1"
    Application.Run strLastDiaBox
End Sub
Sub exp02(control As IRibbonControl)
    strLastDiaBox = "'Exp2.xlam'!showdoe2"
    Application.Run strLastDiaBox
End Sub
Sub exp03(control As IRibbonControl)
    strLastDiaBox = "'Exp2.xlam'!showdoe3"
    Application.Run strLastDiaBox
End Sub
Sub exp04(control As IRibbonControl)
    strLastDiaBox = "'Exp2.xlam'!showdoe4"
    Application.Run strLastDiaBox
End Sub
Sub exp05(control As IRibbonControl)
    strLastDiaBox = "'Exp2.xlam'!showdoe5"
    Application.Run strLastDiaBox
End Sub
'==================Qua=====================
Sub qua01(control As IRibbonControl)        '������׷�
    strLastDiaBox = "'Grap.xlam'!Showhistogram"
    Application.Run strLastDiaBox
End Sub
Sub qua02(control As IRibbonControl)        '�ķ���
    strLastDiaBox = "'Grap.xlam'!ShowParretochart"
    Application.Run strLastDiaBox
End Sub
Sub qua03(control As IRibbonControl)        '������
    strLastDiaBox = "'Grap.xlam'!Showscatter"
    Application.Run strLastDiaBox
End Sub
Sub qua04(control As IRibbonControl)        'x- bar r
    strLastDiaBox = "'Qua.xlam'!Showxbarr"
    Application.Run strLastDiaBox
End Sub
Sub qua05(control As IRibbonControl)        'x - bar s
    strLastDiaBox = "'Qua.xlam'!Showxbars"
    Application.Run strLastDiaBox
End Sub
Sub qua06(control As IRibbonControl)        'I-MR
    strLastDiaBox = "'Qua.xlam'!ShowIMR"
    Application.Run strLastDiaBox
End Sub
Sub qua07(control As IRibbonControl)        'P
    strLastDiaBox = "'Qua.xlam'!Showspcp"
    Application.Run strLastDiaBox
End Sub
Sub qua08(control As IRibbonControl)        'NP
    strLastDiaBox = "'Qua.xlam'!Showspcnp"
    Application.Run strLastDiaBox
End Sub
Sub qua09(control As IRibbonControl)        'C
    strLastDiaBox = "'Qua.xlam'!Showspcc"
    Application.Run strLastDiaBox
End Sub
Sub qua10(control As IRibbonControl)        'U
    strLastDiaBox = "'Qua.xlam'!Showspcu"
    Application.Run strLastDiaBox
End Sub
Sub qua11(control As IRibbonControl)        '���Լ�����
    strLastDiaBox = "'Qua.xlam'!Shownorm"
    Application.Run strLastDiaBox
End Sub
Sub qua12(control As IRibbonControl)        '���Ժ��� �����ɷºм�
    strLastDiaBox = "'Qua.xlam'!Showmnod"
    Application.Run strLastDiaBox
End Sub
Sub qua13(control As IRibbonControl)        '���׺���
    strLastDiaBox = "'Qua.xlam'!Showbino"
    Application.Run strLastDiaBox
End Sub
Sub qua14(control As IRibbonControl)        '���Ƽۺ���
    strLastDiaBox = "'Qua.xlam'!Showpoisson"
    Application.Run strLastDiaBox
End Sub
'============Data mining===========
Sub DM01(control As IRibbonControl)
    strLastDiaBox = "'Dm.xlam'!ShowDm01"
    Application.Run strLastDiaBox
End Sub
Sub DM02(control As IRibbonControl)
    strLastDiaBox = "'Dm.xlam'!ShowDm02"
    Application.Run strLastDiaBox
End Sub
Sub DM03(control As IRibbonControl)
    strLastDiaBox = "'Dm.xlam'!ShowDm03"
    Application.Run strLastDiaBox
End Sub
Sub DM04(control As IRibbonControl)
    strLastDiaBox = "'Dm.xlam'!ShowDm04" '������Ȱ
    Application.Run strLastDiaBox
End Sub
Sub DM05(control As IRibbonControl)
    strLastDiaBox = "'Dm.xlam'!ShowDm05" '�Ƹ���
    Application.Run strLastDiaBox
End Sub
Sub DM06(control As IRibbonControl)
    strLastDiaBox = "'Dm.xlam'!ShowDm06"
    Application.Run strLastDiaBox
End Sub
Sub DM07(control As IRibbonControl)
    strLastDiaBox = "'Dm.xlam'!DmMain"
    Application.Run strLastDiaBox
End Sub
Sub DM08(control As IRibbonControl)
    strLastDiaBox = "'Dm.xlam'!DmMain"
    Application.Run strLastDiaBox
End Sub
'================gene==============
Sub gene00(control As IRibbonControl)
    strLastDiaBox = "'StatGene.xlam'!Genetest"
    Application.Run strLastDiaBox
End Sub
Sub gene01(control As IRibbonControl)
    strLastDiaBox = "'StatGene.xlam'!hypo1"
    Application.Run strLastDiaBox
End Sub
Sub gene02(control As IRibbonControl)
    strLastDiaBox = "'RegGene.xlam'!ShowRehypo1"
    Application.Run strLastDiaBox
End Sub
Sub gene04(control As IRibbonControl)
    strLastDiaBox = "'QuaGene.xlam'!Quahypo"
    Application.Run strLastDiaBox
End Sub

'===============���ã��=============
'Sub fav(control As IRibbonControl)
'    strLastDiaBox = "'DATA.xlam'!ShowSplit"
'    Application.Run strLastDiaBox
'End Sub
'===============�ΰ�=================
Sub logo(control As IRibbonControl)
   Hist_LOGO.Show
   packinstall.packinstall
   

    
End Sub
