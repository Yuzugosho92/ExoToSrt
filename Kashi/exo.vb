Public Class Exo

    'フレームレート
    Private _FrameRate As Decimal
    Public Property FrameRate As Decimal
        Get
            Return _FrameRate
        End Get
        Set(value As Decimal)
            _FrameRate = value
        End Set
    End Property

    'フレームレートスケール
    Private _Scale As Long = 1
    Public Property Scale As Long
        Get
            Return _Scale
        End Get
        Set(value As Long)
            _Scale = value
        End Set
    End Property

    '字幕リスト
    Property _Telop As New List(Of Telop)
    Public ReadOnly Property Telop As List(Of Telop)
        Get
            Return _Telop
        End Get
    End Property

End Class


<Serializable>
Public Class Telop

    Sub New(Rate As Decimal, Svale As Long)
        _FrameRate = Rate
        _Scale = Scale
    End Sub

    'フレームレート
    Private _FrameRate As Decimal
    Public Property FrameRate As Decimal
        Get
            Return _FrameRate
        End Get
        Set(value As Decimal)
            _FrameRate = value
        End Set
    End Property

    'フレームレートスケール
    Private _Scale As Long = 1
    Public Property Scale As Long
        Get
            Return _Scale
        End Get
        Set(value As Long)
            _Scale = value
        End Set
    End Property


    '開始フレーム
    Private _StartFrame As Long
    Public Property StartFrame() As Long
        Get
            Return _StartFrame
        End Get
        Set(value As Long)
            _StartFrame = value
        End Set
    End Property


    '終了フレーム
    Private _EndFrame As Long
    Public Property EndFrame() As Long
        Get
            Return _EndFrame
        End Get
        Set(value As Long)
            _EndFrame = value
        End Set
    End Property


    'フレーム位置から時間を算出
    Private Function FrameToTime(Frame As ULong) As String

        Dim FrameRate As Decimal = _FrameRate / _Scale

        Dim Hour As Long = Frame \ FrameRate \ 3600
        Dim Minutes As Long = (Frame \ FrameRate \ 60) - Hour * 60
        Dim Seconds As Long = (Frame \ FrameRate) - Hour * 3600 - Minutes * 60
        Dim Milliseconds As Long = CInt(Int((Frame Mod FrameRate) / FrameRate * 1000))

        Return Hour.ToString("00") & ":" & Minutes.ToString("00") & ":" & Seconds.ToString("00") & "," & Milliseconds.ToString("000")

    End Function


    '開始時間
    Public ReadOnly Property StartTime() As String
        Get
            Return FrameToTime(StartFrame)
        End Get
    End Property


    '終了時間
    Public ReadOnly Property EndTime() As String
        Get
            Return FrameToTime(EndFrame)
        End Get
    End Property


    '歌詞文字列
    Private _Text As String
    Public Property Text() As String
        Get
            Return _Text
        End Get
        Set(value As String)
            _Text = value
        End Set
    End Property


    '16進法文字列から歌詞を登録
    Public Sub SetHexText(tx As String)

        _Text = HexConverter(tx)

    End Sub


    '16進法文字列を文字列に変換して返す関数
    Private Function HexConverter(Tx As String) As String

        Dim Result As New List(Of Byte)

        '4文字ずつ読み取る
        For i As Integer = 0 To (Tx.Length / 4)
            Dim Str1, Str2 As String
            Str1 = Mid(Tx, 4 * i + 1, 2)
            Str2 = Mid(Tx, 4 * i + 3, 2)

            'もし"0000"なら、文末とみなし終了
            If Str1 = "00" AndAlso Str2 = "00" Then
                Exit For
            End If

            'バイトリストに追加
            Result.Add(Convert.ToByte(Str1, 16))
            Result.Add(Convert.ToByte(Str2, 16))
        Next

        '文字列に変換して返す
        Dim Lyrics As String = System.Text.Encoding.Unicode.GetString(Result.ToArray)

        Lyrics = Lyrics.Replace(vbCrLf, "||")

        Return Lyrics

    End Function

End Class



