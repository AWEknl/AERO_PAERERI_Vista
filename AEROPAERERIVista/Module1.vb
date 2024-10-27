Module Module1
    Sub Main()
        Dim Msg
        Dim ri
        Dim title
        Dim st1
        Dim st2
        Dim st3
        Dim st4
        Dim st5
        Dim helps
        Dim text
        Dim ms1, ms2, ms3
        Dim ms4
        Dim TheDate As Date
        Dim s1, s2, s3, s4, s5
        TheDate = #11/18/2022 5:32:00 PM#
        title = "AERO_PAERERI_Vista"
        st1 = vbQuestion
        st2 = vbInformation
        st3 = vbRetryCancel
        st4 = vbExclamation
        st5 = vbCritical
        helps = 1
        text = 2
        s1 = "你应该想想，从我们第一个项目开始，过了多少天了"
        s2 = "是的，已经过去这么多天了，可我们却也从未改变了"
        s3 = "自己对彩蛋的看法"
        s4 = "End By AERO_PAERERI_Vista."
        s5 = "Author by......but I don't know."
        ms1 = MsgBox(s1, st1, title)
        Msg = "可已经过去了" & DateDiff("d", Now, TheDate) & "-天"
        ms2 = MsgBox(Msg, st2, title)
        ms4 = MsgBox(s2, st2, title)
        ri = MsgBox(s3, st3, title)
        If ri = vbRetry Then
            ms3 = MsgBox(s4, st4, title)
        Else
            ms3 = MsgBox(s5, st5, title)
            'ppt中其他功能
            MsgBox("ppt中其他功能")
            MsgBox("弹出音量合成器，已废弃")
            Shell("SndVol.exe shell32.dll,Control_RunDLL mmsys.cpl,,0") '弹出音量合成器
            Dim wshShell
            wshShell = CreateObject("Wscript.Shell") '如果使用VBA需改成' Set Ws = CreateObject("Wscript.Shell")
            MsgBox("silence音量")
            wshShell.SendKeys("棴") 'silence音量
            MsgBox("add音量")
            wshShell.SendKeys("棷") 'add音量
            MsgBox("subtract音量")
            wshShell.SendKeys("棶") 'subtract音量

        End If

    End Sub
    Private Sub CESH_Click() 'old
        Dim Msg
        Dim TheDate As Date
        TheDate = #3/15/2023 5:32:00 PM#
        MsgBox("你应该想想，从我们第一个项目开始，过了多少天了")
        Msg = "---------------" & DateDiff("d", Now, TheDate) & "day---------------"
        MsgBox(Msg)
        MsgBox("是的，已经过去这么多天了，可我们却也从未改变了")
        MsgBox("自己对彩蛋的看法")
    End Sub
End Module
