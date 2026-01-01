Imports System.Runtime.InteropServices
Imports System.Text

Public Class Form1

    Dim FileAddress As String
    Dim FileType As Type

    Private Sub TextBox1_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles TextBox1.DragEnter
        'コントロール内にドラッグされたとき実行される
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            'ドラッグされたデータ形式を調べ、ファイルのときはコピーとする
            e.Effect = DragDropEffects.Copy
        Else
            'ファイル以外は受け付けない
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub TextBox1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles TextBox1.DragDrop
        'コントロール内にドロップされたとき実行される
        'ドロップされたすべてのファイル名を取得する
        Dim fileName As String() = CType(e.Data.GetData(DataFormats.FileDrop, False), String())
        'TextBoxに追加する
        TextBox1.Text = fileName(0)
    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        FIleAddress = TextBox1.Text

        '拡張子からAviUtlのバージョンを選択
        Select Case IO.Path.GetExtension(FileAddress).ToLower
            Case ".exo"
                FileType = Type.v1
            Case ".scene"
                FileType = Type.v2
            Case Else
                Exit Sub
        End Select

        Dim Reader As IO.StreamReader


        If FileType = Type.v1 Then
            'AviUtl1ならば
            'シフトJISを使うためのおまじない
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)
            'ファイルを開く
            Reader = New IO.StreamReader(FileAddress, System.Text.Encoding.GetEncoding("Shift_JIS"))

        Else
            'AviUtl2ならば
            'ファイルを開く
            Reader = New IO.StreamReader(FileAddress, System.Text.Encoding.UTF8)
        End If

        '1行読み込み
        Dim Strs As String()
        Dim Tx As String = Reader.ReadLine


        If FileType = Type.v1 AndAlso Tx = "[exedit]" Then
            '処理を続行
        ElseIf FileType = Type.v2 AndAlso Tx = "[scene]" Then
            '処理を続行
        Else
            '[exedit]が見つからなければ、ここで終了
            Reader.Close()
            Exit Sub
        End If

        Dim Exo As New exo

        Do
            '1行読み込む
            Tx = Reader.ReadLine

            Strs = Strings.Split(Tx, "=")

            Select Case Strs(0)
                Case "rate", "video.rate"
                    'フレームレートを登録
                    Exo.FrameRate = CDec(Strs(1))

                Case "scale", "video.scale"
                    'フレームスケールを登録
                    Exo.Scale = CLng(Strs(1))
                    Exit Do
            End Select

        Loop



        Dim oTelop As New Telop(Exo.FrameRate, Exo.Scale)

        Do
            '1行読み込む
            Tx = Reader.ReadLine

            If Tx Is Nothing Then
                '文末まで読み込んだらDoを出る
                Reader.Close()
                Exit Do
            End If

            Strs = Split(Tx, "=")

            If FileType = Type.v1 Then
                'AviUtl1の場合
                Select Case Strs(0)
                    Case "start"
                        '新しいテロップクラスを作成
                        oTelop = New Telop(Exo.FrameRate, Exo.Scale)

                        oTelop.StartFrame = CULng(Strs(1))

                    Case "end"
                        oTelop.EndFrame = CULng(Strs(1))

                    Case "_name"

                        If Strs(1) = "テキスト" Then

                            Do
                                '1行読み込む
                                Tx = Reader.ReadLine

                                Strs = Strings.Split(Tx, "="c)

                                If Strs(0) = "text" Then


                                    oTelop.SetHexText(Strs(1))



                                    If Not oTelop.Text = "[none]" Then


                                        Exo.Telop.Add(oTelop)
                                    End If




                                    Exit Do

                                End If
                            Loop





                        End If
                End Select

            Else

                'AviUtl2の場合
                Select Case Strs(0)
                    Case "frame"
                        '新しいテロップクラスを作成
                        oTelop = New Telop(Exo.FrameRate, Exo.Scale)

                        Dim oFrame = Strs(1).Split(","c)

                        '開始位置を登録
                        oTelop.StartFrame = CLng(oFrame(0))
                        '終了位置を登録
                        oTelop.EndFrame = CLng(oFrame(1))

                    Case "effect.name"

                        If Strs(1) = "テキスト" Then

                            Do
                                '1行読み込む
                                Tx = Reader.ReadLine

                                Strs = Strings.Split(Tx, "=")

                                If Strs(0) = "テキスト" Then



                                    For i As Integer = 1 To Strs.Length - 1

                                        If i > 1 Then
                                            oTelop.Text &= "="
                                        End If

                                        oTelop.Text &= Strs(i)
                                    Next




                                    If Not oTelop.Text = "[none]" Then


                                        Exo.Telop.Add(oTelop)
                                    End If

                                    'Doから出る
                                    Exit Do
                                End If
                            Loop

                        End If
                End Select

            End If




        Loop





        Dim Str As New System.Text.StringBuilder

        For i As Integer = 0 To Exo.Telop.Count - 1


            Str.AppendLine(i + 1.ToString)

            Str.Append(Exo.Telop(i).StartTime & " --> ")

            If i < Exo.Telop.Count - 1 Then

                If Exo.Telop(i + 1).StartFrame - Exo.Telop(i).EndFrame <= Exo.FrameRate Then

                    Str.AppendLine(Exo.Telop(i + 1).StartTime)

                Else

                    Str.AppendLine(Exo.Telop(i).EndTime)


                End If


            End If



            Str.AppendLine(Exo.Telop(i).Text & vbCrLf)

        Next







        'srtファイルを書き込む
        Dim Writer As New IO.StreamWriter(IO.Path.GetDirectoryName(FileAddress) & "\" & IO.Path.GetFileNameWithoutExtension(FileAddress) & ".srt", False, System.Text.Encoding.UTF8)

        Writer.WriteLine(Str.ToString)

        Writer.Close()

    End Sub






    Public Enum Type
        v1
        v2
    End Enum


End Class






