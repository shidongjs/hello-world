# hello-world
Private Sub CommandButton1_Click()

'编程：石东，电话：13912253120，电邮：shidong@139.com

    Dim myFilekg As String
    Dim myFilewg As String
    Dim myFilejs As String
    Dim myFileys As String
    
    Dim docApp As Object
    
    myFilekg = ThisWorkbook.Path & "\kgbg.doc"
    myFilewg = ThisWorkbook.Path & "\wgbg.doc"
    myFilejs = ThisWorkbook.Path & "\wgjs.doc"
    myFileys = ThisWorkbook.Path & "\wgys.doc"
    Dim xmbh, xmmc, jhkg, jhwg, kgbg, tbrq As String
    
    Dim i As Integer '表格行号
    Dim n As Integer '实际生成项目计数
    n = 0
    
    i = 2
    
     Dim FinalRow As Integer '表格最后一行
    
    FinalRow = Sheets("Sheet1").Range("B65536").End(xlUp).Row
    
    Set docApp = CreateObject("Word.Application")
    
    For i = 2 To FinalRow
    
         xmbh = Worksheets("sheet1").Range("a" & i) '项目编号
         xmmc = Worksheets("sheet1").Range("b" & i) '项目名称
         jhkg = Format(Worksheets("sheet1").Range("c" & i).Value, "yyyy" & """" & "年" & """" & "m" & """" & "月" & """" & "d" & """" & "日" & """" & "") '计划开工
         jhwg = Format(Worksheets("sheet1").Range("d" & i).Value, "yyyy" & """" & "年" & """" & "m" & """" & "月" & """" & "d" & """" & "日" & """" & "") '计划完工
         kgbg = Format(Worksheets("sheet1").Range("f" & i).Value, "yyyy" & """" & "年" & """" & "m" & """" & "月" & """" & "d" & """" & "日" & """" & "") '开工报告
         kgtbrq = Format(Worksheets("sheet1").Range("g" & i).Value, "yyyy" & """" & "年" & """" & "m" & """" & "月" & """" & "d" & """" & "日" & """" & "") '开工填报日期
         wgtbrq = Format(Worksheets("sheet1").Range("h" & i).Value, "yyyy" & """" & "年" & """" & "m" & """" & "月" & """" & "d" & """" & "日" & """" & "") '完工填报日期
        
       
         MkDir ThisWorkbook.Path & "\" & xmmc                       '创建项目名称文件夹
            
               With docApp
                   .documents.Open myFilekg
                   .Visible = True
                   .activedocument.Tables.Item(1).Cell(2, 4).Range.Text = xmbh
                   .activedocument.Tables.Item(1).Cell(3, 2).Range.Text = xmmc
                   .activedocument.Tables.Item(1).Cell(9, 2).Range.Text = jhkg
                   .activedocument.Tables.Item(1).Cell(9, 4).Range.Text = jhwg
                   .activedocument.Tables.Item(1).Cell(17, 1).Range.Text = "本工程已于" & kgbg & "正式开工，特此报告。"
                   .activedocument.Tables.Item(1).Cell(23, 1).Range.Text = "  填报日期：" & kgtbrq
                   .activedocument.SaveAs ThisWorkbook.Path & "\" & xmmc & "\" & xmbh & "-KG" & Format(Worksheets("sheet1").Range("f" & i).Value, "YYYYMMDD") & "-" & xmmc & ".doc"
                   .activedocument.Close
               End With
               
               With docApp
                   .documents.Open myFilewg
                   .Visible = True
                   .activedocument.Tables.Item(1).Cell(2, 4).Range.Text = xmbh
                   .activedocument.Tables.Item(1).Cell(3, 2).Range.Text = xmmc
                   .activedocument.Tables.Item(1).Cell(8, 2).Range.Text = jhkg
                   .activedocument.Tables.Item(1).Cell(8, 4).Range.Text = jhwg
                   .activedocument.Tables.Item(1).Cell(10, 1).Range.Text = "已完成江苏移动南通地区" & xmmc & "。"
                   .activedocument.Tables.Item(1).Cell(19, 1).Range.Text = "填报日期：" & wgtbrq
                   .activedocument.SaveAs ThisWorkbook.Path & "\" & xmmc & "\" & xmbh & "-WG" & Format(Worksheets("sheet1").Range("h" & i).Value, "YYYYMMDD") & "-" & xmmc & ".doc"
                    .activedocument.Close
               End With
               
                With docApp
                   .documents.Open myFilejs
                   .Visible = True
                   .activedocument.SaveAs ThisWorkbook.Path & "\" & xmmc & "\" & xmbh & "-JS" & Format(Worksheets("sheet1").Range("E" & i).Value, "YYYYMMDD") & "-" & xmmc & ".doc"  '空白结算word
                   .activedocument.Close
                End With
    
                With docApp
                   .documents.Open myFileys
                   .Visible = True
                   .activedocument.SaveAs ThisWorkbook.Path & "\" & xmmc & "\" & xmbh & "-YS" & Format(Worksheets("sheet1").Range("I" & i).Value, "YYYYMMDD") & "-" & xmmc & ".doc"  '空白验收word
                   .activedocument.Close
                End With
    
    
    
        n = n + 1
        
    
   Next i
   
   docApp.Quit
    Set docApp = Nothing


    MsgBox (n & "个文件夹及相应文件创建完成!")





End Sub
