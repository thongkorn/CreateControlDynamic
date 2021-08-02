' / ---------------------------------------------------------------
' / Developer : Mr.Surapon Yodsanga (Thongkorn Tubtimkrob)
' / eMail : thongkorn@hotmail.com
' / URL: http://www.g2gnet.com (Khon Kaen - Thailand)
' / Facebook: https://www.facebook.com/g2gnet (For Thailand)
' / Facebook: https://www.facebook.com/commonindy (Worldwide)
' / More Info: http://www.g2gnet.com/webboard
' /
' / Purpose: Create controls dynamic.
' / Microsoft Visual Basic .NET (2010) + MS Access 2010+
' /
' / This is open source code under @CopyLeft by Thongkorn Tubtimkrob.
' / You can modify and/or distribute without to inform the developer.
' / ---------------------------------------------------------------
Imports System.Data.OleDb

Public Class frmMainShow
    '// Start Here
    Private Sub frmMainShow_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Conn = MyDBModule()
        Call CreateTabPagePanel()
    End Sub

    Private Sub frmMainShow_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Result As Byte = MessageBox.Show("คุณแน่ใจว่าต้องการจบการทำงานของโปรแกรม?", "ยืนยันการปิดโปรแกรม", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
        If Result = DialogResult.Yes Then
            Me.Dispose()
            If Conn.State = ConnectionState.Open Then Conn.Close()
            GC.SuppressFinalize(Me)
            Application.Exit()
        Else
            e.Cancel = True
        End If
    End Sub

    ' / ---------------------------------------------------------------
    ' / สร้างกลุ่ม Control ประกอบด้วย TabPage โดยนำเอา Panel แสดงในบน TabPage
    ' / และในแต่ละ Panel ให้แสดงผลปุ่มคำสั่ง
    ' / ---------------------------------------------------------------
    Private Sub CreateTabPagePanel()
        ' / ค้นหาจำนวนของกลุ่มสินค้าเข้ามาก่อน เพื่อกำหนดจำนวนของ TabPages ได้
        Dim CategoryCount As Byte = 0
        strSQL = _
            "SELECT Category.CategoryPK, Category.CategoryName " & _
            " FROM(Category) " & _
            " ORDER BY Category.CategoryPK "
        If Conn.State = ConnectionState.Closed Then Conn.Open()
        Cmd = New OleDb.OleDbCommand(strSQL, Conn)
        DR = Cmd.ExecuteReader
        '// เก็บค่า CategoryName เพื่อแสดงผลชื่อใน TabPages
        Dim ListCat As List(Of String) = New List(Of String)()
        If DR.HasRows Then
            While DR.Read()
                '// เก็บค่าเพื่อแสดงชื่อบน TabPage
                ListCat.Add(DR.Item("CategoryName").ToString)
                CategoryCount += 1
            End While
        End If
        DR.Close()
        Cmd.Dispose()
        ' / ใส่จำนวน Panel ลงบน TabControl
        Dim pn(CategoryCount) As Panel
        Dim iCat As Byte = 0
        For i = 0 To CategoryCount - 1
            ' / Create a tabpage
            Dim tabPageRef As New TabPage
            ' / Set the tabpage to be your desired tab
            tabPageRef.Name = "Tab" & CStr(i)
            '// แสดงชื่อกลุ่ม Category
            tabPageRef.Text = ListCat(iCat).ToString
            '// เพิ่ม TabPage ลงใน TabControl1
            Me.TabControl1.Controls.Add(tabPageRef)
            tabPageRef = TabControl1.TabPages(i)
            '// Create Panel
            pn(i) = New Panel
            pn(i).Name = "pn" & i
            With pn(i)
                .Location = New System.Drawing.Point(1, 1)
                .Size = New System.Drawing.Size(TabControl1.Width - 10, TabControl1.Height - 30)
                .BackColor = Color.Moccasin
                .AutoScroll = True
                .Anchor = AnchorStyles.Bottom + AnchorStyles.Top
            End With
            '/ Add the panel into the TabControl
            tabPageRef.Controls.Add(pn(i))
            strSQL = _
                " SELECT Food.FoodPK, Food.FoodID, Food.FoodName, Food.PriceCash, Food.PictureFood, Food.CategoryFK" & _
                " FROM(Food) " & _
                " WHERE Food.CategoryFK = " & i + 1 & _
                " ORDER BY FoodPK "
            '// สร้างปุ่ม (Button) ลงใน Panel (จำนวนหลัก, รายการ, Panel) ... ตัวอย่างนี้ผมใช้ 4 หลัก
            Call CreateButtons(4, strSQL, pn(i))
            iCat += 1
        Next

    End Sub

    ' / ---------------------------------------------------------------
    '// เพิ่มปุ่มคำสั่ง (Button Control) แบบ Run Time 
    Private Sub CreateButtons(ByVal ColCount As Byte, ByVal sql As String, pn As Panel)
        Dim Buttons As New Dictionary(Of String, Button)
        Dim LBs As New Dictionary(Of String, Label)
        Dim img As String
        Dim Rec As Integer = 0
        Try
            If Conn.State = ConnectionState.Closed Then Conn.Open()
            Cmd = New OleDbCommand(sql, Conn)
            DR = Cmd.ExecuteReader
            ' / Make sure Primary Key only one and not duplicate
            While DR.Read()
                If DR.HasRows Then
                    Dim B As New Button
                    Dim LB As New Label
                    pn.Controls.Add(B)
                    pn.Controls.Add(LB)
                    LB.Height = 32
                    With B
                        '// กำหนดขนาด
                        .Height = 140
                        .Width = 140
                        '// หาตำแหน่งซ้าย (Left) ... ด้วยการหารเอาเศษ (Mod)
                        '// การหารเอาเศษจะได้ค่าสูงสุด คือ ค่าตัวหาร (Mod) ลบออก 1 เช่น X Mod 4 จะได้ค่า 0, 1, 2, 3
                        '// ทำให้รู้ตำแหน่งหลัก
                        '// เอาค่าที่ได้ในแต่ละหลักมาคูณความกว้างของปุ่มคำสั่ง
                        .Left = (Rec Mod ColCount) * B.Width
                        '// หาตำแหน่งบน (Top)
                        '// แถวแรกคือ 0 ก็ให้ Top อยู่เท่าเดิม
                        If Rec = 0 Then
                            .Top = (Rec \ ColCount) * B.Height
                            '// แถวถัดไปเอาค่าที่ได้จากการหารตัดเศษในแต่ละแถว X (ความสูงของปุ่ม + ความสูงของลาเบล)
                        Else
                            .Top = (Rec \ ColCount) * (B.Height + LB.Height)
                        End If
                        '//
                        .Text = DR.Item("FoodName")
                        '// นำค่า Primary Key ไปเก็บไว้ที่คุณสมบัติ Tag เมื่อกดปุ่มคำสั่งจะใช้ค่านี้ไปค้นหาจากฐานข้อมูล
                        .Tag = DR.Item("FoodPK")
                        '// อ่านค่ารูปภาพ และตรวจสอบการมีอยู่จริงของภาพด้วยฟังค์ชั่น GetImages
                        img = GetImages(DR.Item("PictureFood"))
                        '// ใส่ภาพลงไปในปุ่มคำสั่ง Button
                        .BackgroundImage = New System.Drawing.Bitmap(img)
                        .Cursor = Cursors.Hand
                        .BackgroundImageLayout = ImageLayout.Stretch
                        .Font = New Font("Century Gothic", 10, FontStyle.Bold)
                        If pn.Name = "pn2" Then
                            .ForeColor = Color.Black
                        Else
                            .ForeColor = Color.LightYellow
                        End If
                        .TextImageRelation = TextImageRelation.ImageAboveText
                        .TextAlign = ContentAlignment.BottomCenter
                        .UseVisualStyleBackColor = True
                    End With
                    '// ใส่ปุ่มคำสั่ง Button เข้าไป
                    Buttons.Add(B.Text, B)
                    '// Label
                    With LB
                        .AutoSize = False
                        .Width = B.Width - 2
                        .Left = B.Left + 1
                        If Rec <= (ColCount - 1) Then
                            .Top = B.Height
                        Else
                            .Top = B.Height + B.Top
                        End If
                        '.ForeColor = Color.Black
                        '.BackColor = Color.DeepSkyBlue
                        If Rec Mod 2 = 0 Then
                            .BackColor = Color.Orange
                        Else
                            .BackColor = Color.Crimson
                        End If
                        .TextAlign = ContentAlignment.MiddleCenter
                        .Text = DR.Item("FoodName") & vbCrLf & Format(CDbl(DR.Item("PriceCash")), "#,##0.00") & "B."
                        .Font = New Font("Tahoma", 9, FontStyle.Bold)
                    End With
                    LBs.Add(B.Text, LB)
                    '//
                    Rec += 1
                    '// Force events handler.
                    AddHandler B.Click, AddressOf ClickButton
                End If
            End While
            DR.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    '// Click Button event, get the text of button
    Public Sub ClickButton(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim btn As Button = sender
        MessageBox.Show("คุณคลิ๊ก [" + btn.Text + "]" & vbCrLf & "Tag=" & btn.Tag & " จะเป็นค่า Primary Key" & vbCrLf & "เพื่อไป Query ค้นหาข้อมูลรายละเอียดทั้งหมด.")
    End Sub

    '// หากตำแหน่งไฟล์ภาพ
    Private Function GetImages(ByVal picData As String) As String
        '// Show picture in cell.
        If picData <> "" Then
            '// First, before load data into DataGrid and check File exists or not?
            If Dir(strPathImages & picData) = "" Then
                GetImages = strPathImages & "FoodStuff.png"
            Else
                GetImages = strPathImages & picData
            End If

            '// ไม่มีข้อมูลภาพ
        Else
            GetImages = strPathImages & "FoodStuff.png"
        End If
    End Function

End Class
