Public Class MainWnd

  Private m_imgSrc As Image
  Private m_imgPnl(4, 4) As Image

  Private m_nAcross As Integer
  Private m_nDown As Integer

  Private Const nLen As Integer = 14 ' len  of line
  Private Const nEdge As Integer = 5 ' contrasting edge width

  Private Sub MainWnd_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
  End Sub


  Private Sub btnFile_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFile.Click
    Dim d As New OpenFileDialog

    If d.ShowDialog() <> Windows.Forms.DialogResult.OK Then Return

    txtFile.Text = d.FileName
  End Sub

  Private Sub btnProcess_click(ByVal sender As Object, ByVal e As EventArgs) Handles btnProcess.Click
    Dim g As Graphics
    Dim x, y, i, k, szw, szh As Integer
    Dim w, h As Integer
    Dim oPenYF As New Pen(Color.FromArgb(&H90, &H90, &H90), 3)
    Dim oBmp As Image

    If m_nAcross = 0 Then
      MsgBox("Choose number of printed sheets across", MsgBoxStyle.Information)
      Return
    End If

    If m_nDown = 0 Then
      MsgBox("Choose number of printed sheets down", MsgBoxStyle.Information)
      Return
    End If

    If tabCtrl.SelectedIndex = 0 Then '0 = portrait
      txtFeet.Text = (13 * m_nAcross) \ 12 & "' " & (13 * m_nAcross) Mod 12 & """ X " & (19 * m_nDown) \ 12 & "' " & (19 * m_nDown) Mod 12 & """"
      txtInch.Text = 13 * m_nAcross & """ X " & 19 * m_nDown & """"
    Else
      txtFeet.Text = (19 * m_nAcross) \ 12 & "' " & (19 * m_nAcross) Mod 12 & """ X " & (13 * m_nDown) \ 12 & "' " & (13 * m_nDown) Mod 12 & """"
      txtInch.Text = 19 * m_nAcross & """ X " & 13 * m_nDown & """"
    End If

    ' final images

    ' background
    ' picture
    ' borders
    ' alignment lines

    oBmp = New Bitmap(txtFile.Text)
    m_imgSrc = New Bitmap(oBmp.Width + 12, oBmp.Height + 12)
    g = Graphics.FromImage(m_imgSrc)
    g.PageUnit = GraphicsUnit.Pixel
    g.FillRectangle(Brushes.Black, 0, 0, m_imgSrc.Width, m_imgSrc.Height)
    g.DrawImage(oBmp, New Rectangle(6, 6, oBmp.Width, oBmp.Height), New Rectangle(0, 0, oBmp.Width, oBmp.Height), GraphicsUnit.Pixel) ' this might not be doing what is expected

    ' m_imgSrc.Save("e:\pictures\test.jpg", Drawing.Imaging.ImageFormat.Jpeg)

    'm_imgSrc = New Bitmap(txtFile.Text) oops

    w = m_imgSrc.Width / m_nAcross ' including border
    h = m_imgSrc.Height / m_nDown

    szw = w + 5
    szh = h + 5

    For i = 0 To 4
      For k = 0 To 4
        If Not m_imgPnl(i, k) Is Nothing Then
          m_imgPnl(i, k).Dispose()
          m_imgPnl(i, k) = Nothing
        End If
      Next
    Next

    For i = 0 To m_nDown - 1 ' down
      For k = 0 To m_nAcross - 1 ' across

        x = w * k
        y = h * i
        m_imgPnl(i, k) = New Bitmap(w + 5, h + 5, Imaging.PixelFormat.Format32bppRgb)
        g = Graphics.FromImage(m_imgPnl(i, k))
        '  g.DrawImage(m_imgSrc, 0, 0, New Rectangle(x, y, w, h), GraphicsUnit.Pixel)
        g.DrawImage(m_imgSrc, New Rectangle(0, 0, w, h), New Rectangle(x, y, w, h), GraphicsUnit.Pixel)

        If Not (i = 0 And k = 0) Then DrawAlignUpperLeft(g, szw, szh) ' alignment guide line
        If Not (i = 0 And k = m_nAcross - 1) Then DrawAlignUpperRight(g, szw, szh) ' alignment guide line
        If Not (i = m_nDown - 1 And k = 0) Then DrawAlignLowerLeft(g, szw, szh) ' alignment guide line
        If Not (i = m_nDown - 1 And k = m_nAcross - 1) Then DrawAlignLowerRight(g, szw, szh) ' alignment guide line 

        g.Dispose()

      Next
    Next
    ' visables

    If tabCtrl.SelectedIndex = 0 Then

      img_P00.Image = m_imgPnl(0, 0)
      img_P01.Image = m_imgPnl(0, 1)
      img_P02.Image = m_imgPnl(0, 2)
      img_P03.Image = m_imgPnl(0, 3)
      img_P04.Image = m_imgPnl(0, 4)

      img_P10.Image = m_imgPnl(1, 0)
      img_P11.Image = m_imgPnl(1, 1)
      img_P12.Image = m_imgPnl(1, 2)
      img_P13.Image = m_imgPnl(1, 3)
      img_P14.Image = m_imgPnl(1, 4)

      img_P20.Image = m_imgPnl(2, 0)
      img_P21.Image = m_imgPnl(2, 1)
      img_P22.Image = m_imgPnl(2, 2)
      img_P23.Image = m_imgPnl(2, 3)
      img_P24.Image = m_imgPnl(2, 4)

      img_P30.Image = m_imgPnl(3, 0)
      img_P31.Image = m_imgPnl(3, 1)
      img_P32.Image = m_imgPnl(3, 2)
      img_P33.Image = m_imgPnl(3, 3)
      img_P34.Image = m_imgPnl(3, 4)

      img_P40.Image = m_imgPnl(4, 0)
      img_P41.Image = m_imgPnl(4, 1)
      img_P42.Image = m_imgPnl(4, 2)
      img_P43.Image = m_imgPnl(4, 3)
      img_P44.Image = m_imgPnl(4, 4)

    Else

      img_L00.Image = m_imgPnl(0, 0)
      img_L01.Image = m_imgPnl(0, 1)
      img_L02.Image = m_imgPnl(0, 2)
      img_L03.Image = m_imgPnl(0, 3)
      img_L04.Image = m_imgPnl(0, 4)

      img_L10.Image = m_imgPnl(1, 0)
      img_L11.Image = m_imgPnl(1, 1)
      img_L12.Image = m_imgPnl(1, 2)
      img_L13.Image = m_imgPnl(1, 3)
      img_L14.Image = m_imgPnl(1, 4)

      img_L20.Image = m_imgPnl(2, 0)
      img_L21.Image = m_imgPnl(2, 1)
      img_L22.Image = m_imgPnl(2, 2)
      img_L23.Image = m_imgPnl(2, 3)
      img_L24.Image = m_imgPnl(2, 4)

      img_L30.Image = m_imgPnl(3, 0)
      img_L31.Image = m_imgPnl(3, 1)
      img_L32.Image = m_imgPnl(3, 2)
      img_L33.Image = m_imgPnl(3, 3)
      img_L34.Image = m_imgPnl(3, 4)

      img_L40.Image = m_imgPnl(4, 0)
      img_L41.Image = m_imgPnl(4, 1)
      img_L42.Image = m_imgPnl(4, 2)
      img_L43.Image = m_imgPnl(4, 3)
      img_L44.Image = m_imgPnl(4, 4)

    End If


    btnSave.Enabled = True

  End Sub

  '
  '
  '       ptTop
  '        |
  '     x  |
  '  ______|
  'ptLft   ptBot

  Private Sub DrawAlignLowerRight(ByVal g As Graphics, ByVal szw As Integer, ByVal szh As Integer)
    Dim oPen As New Pen(Color.DimGray, 1)
    Dim oPen2 As New Pen(Color.Silver, 1)

    g.DrawLine(oPen, New Point(szw - (1 + nEdge), szh - (nLen + 1 + nEdge)), New Point(szw - (1 + nEdge), szh - (1 + nEdge)))
    g.DrawLine(oPen, New Point(szw - (nLen + 1 + nEdge), szh - (1 + nEdge)), New Point(szw - (2 + nEdge), szh - (1 + nEdge)))

    g.DrawLine(oPen2, New Point(szw - (2 + nEdge), szh - (nLen + 1 + nEdge)), New Point(szw - (2 + nEdge), szh - (2 + nEdge))) ' 1 less
    g.DrawLine(oPen2, New Point(szw - (nLen + 1 + nEdge), szh - (2 + nEdge)), New Point(szw - (2 + nEdge), szh - (2 + nEdge)))


  End Sub

  '
  '
  ' ptTop
  '  |      
  '  |   x   
  '  |______
  ' ptBot   ptRgt

  Private Sub DrawAlignLowerLeft(ByVal g As Graphics, ByVal szw As Integer, ByVal szh As Integer)
    Dim oPen As New Pen(Color.DimGray, 1)
    Dim oPen2 As New Pen(Color.Silver, 1)

    g.DrawLine(oPen, New Point(0, szh - (nLen + 1 + nEdge)), New Point(0, szh - (nEdge + 1)))
    g.DrawLine(oPen, New Point(1, szh - (1 + nEdge)), New Point(1 + nLen, szh - (1 + nEdge)))

    g.DrawLine(oPen2, New Point(1, szh - (nLen + 1 + nEdge)), New Point(1, szh - (nEdge + 2)))
    g.DrawLine(oPen2, New Point(1, szh - (2 + nEdge)), New Point(1 + nLen, szh - (2 + nEdge)))

  End Sub


  '  ______
  '        |
  '     x  |
  '        |
  '
  '
  '
  '

  Private Sub DrawAlignUpperRight(ByVal g As Graphics, ByVal szw As Integer, ByVal szh As Integer)
    Dim oPen As New Pen(Color.DimGray, 1)
    Dim oPen2 As New Pen(Color.Silver, 1)

    g.DrawLine(oPen, New Point(szw - (nLen + 1 + nEdge), 0), New Point(szw - (nEdge + 1), 0))
    g.DrawLine(oPen, New Point(szw - (1 + nEdge), 1), New Point(szw - (1 + nEdge), 1 + nLen))

    g.DrawLine(oPen2, New Point(szw - (nLen + 1 + nEdge), 1), New Point(szw - (nEdge + 2), 1))
    g.DrawLine(oPen2, New Point(szw - (2 + nEdge), 1), New Point(szw - (2 + nEdge), 1 + nLen))

  End Sub

  '   ______
  '  |      
  '  |   x  
  '  |      
  '
  '
  '
  '

  Private Sub DrawAlignUpperLeft(ByVal g As Graphics, ByVal szw As Integer, ByVal szh As Integer)
    Dim oPen As New Pen(Color.DimGray, 1)
    Dim oPen2 As New Pen(Color.Silver, 1)

    g.DrawLine(oPen, New Point(1, 0), New Point(1 + nLen, 0))
    g.DrawLine(oPen, New Point(0, 1), New Point(0, 1 + nLen))

    g.DrawLine(oPen2, New Point(1, 1), New Point(1 + nLen, 1))
    g.DrawLine(oPen2, New Point(1, 1), New Point(1, 1 + nLen))

  End Sub

  Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
    Dim dlg As FolderBrowserDialog
    Dim i, k As Integer
    Dim strName As String = ""
    Dim oDir As IO.DirectoryInfo
    Dim oFile As IO.FileInfo
    Dim strPath As String = ""
    Dim ret As DialogResult

    dlg = New FolderBrowserDialog
    dlg.SelectedPath = GetSetting("Collage", "Settings", "SaveDirectory", "")
    ret = dlg.ShowDialog()
    If ret = DialogResult.OK Then
      strPath = dlg.SelectedPath
    End If
    dlg.Dispose()
    If ret <> DialogResult.OK Then Return

    Try
      IO.Directory.CreateDirectory(strPath)
    Catch
    End Try

    oDir = New IO.DirectoryInfo(strPath)
    For Each oFile In oDir.GetFiles
      Try
        oFile.Delete()
      Catch
      End Try
    Next

    For i = 0 To 2 ' down
      For k = 0 To 2 ' across
        If Not m_imgPnl(i, k) Is Nothing Then

          strName = "Y" & (i + 1) & "_X" & (k + 1)
          Try
            m_imgPnl(i, k).Save(strPath & "\" & strName & ".bmp", System.Drawing.Imaging.ImageFormat.Bmp)
          Catch
            MsgBox("Stopping save process. Unable to save file e:\pictures\collages\Program\" & strName & ".bmp (" & Err.Description & ")", MsgBoxStyle.Information)
            Return
          End Try

        End If
      Next
    Next

    SaveSetting("Collage", "Settings", "SaveDirectory", strPath)

    MsgBox("Done", MsgBoxStyle.Information)

  End Sub

  Private Sub cboAcross_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboAcross.SelectedIndexChanged
    m_nAcross = CInt(cboAcross.SelectedItem.ToString)
  End Sub

  Private Sub cboDown_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboDown.SelectedIndexChanged
    m_nDown = CInt(cboDown.SelectedItem.ToString)
  End Sub

End Class
