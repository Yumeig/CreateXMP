Attribute VB_Name = "CreateXMP"
Sub CreateXMP()

Dim tmpDelayTime As Single
Dim tmpDuration As Single
Dim tmpDuration_up As Single
Dim tmpTypa As Integer
Dim tmpTypa_up As Integer
Dim i As Integer
Dim j As Integer
Dim e As Single '总时长
Dim emax As Single
Dim list(364) As Single
Dim f As Single
Dim template_marker As String
Dim xmp_template As String
Dim file As String


'Main
For i = 1 To ActivePresentation.Slides.Count

    For j = 1 To ActivePresentation.Slides.Item(i).TimeLine.MainSequence.Count
    
        tmpDelayTime = ActivePresentation.Slides.Item(i).TimeLine.MainSequence.Item(j).Timing.TriggerDelayTime
        tmpDuration = tmpDelayTime + ActivePresentation.Slides.Item(i).TimeLine.MainSequence.Item(j).Timing.Duration
        tmpType = ActivePresentation.Slides.Item(i).TimeLine.MainSequence.Item(j).Timing.TriggerType
        
        '上一元素动画时长及类型
        If j <> 1 Then
            tmpDuration_up = ActivePresentation.Slides.Item(i).TimeLine.MainSequence.Item(j - 1).Timing.Duration
            tmpType_up = ActivePresentation.Slides.Item(i).TimeLine.MainSequence.Item(j - 1).Timing.TriggerType
        End If

        '类型2 上一项同时，取动画时间最长的
        If tmpType = 2 Then
            If tmpDuration > tmpDuration_up And tmpDuration > emax Then
                emax = tmpDuration
            End If
            
            If tmpType_up = 1 Then
                e = e - tmpDuration_up + emax
            Else
                e = emax
            End If
        End If
        
        '类型1 点击
        If tmpType = 1 Then
            e = e + tmpDuration
            '重置emax
            emax = 0
        End If
        
        '类型3 上一项后
        If tmpType = 3 Then
            e = e + tmpDuration
        End If
        
        Debug.Print e; emax; tmpDuration; tmpDuration_up
    Next j
    
    '幻灯片时长问题(与创建视频每页时间相关)
    list(i) = e
    f = list(i) - list(i - 1) '求当页动画时长
    If f < 5 Then
        e = e + 5 - f
        Debug.Print "not"; f
    End If
    
    '加入幻灯片切换时间(排除效果“无”)
    If ActivePresentation.Slides(i).SlideShowTransition.EntryEffect <> 0 Then
        e = e + ActivePresentation.Slides(i).SlideShowTransition.Duration
    End If
    
    '变更值
    list(i) = e
    
    '提取单页时间
    Debug.Print "第"; i; "页"; e; "秒"
    
    'ASCII Chr() https://baike.baidu.com/item/Chr/580328
    '字符串末位加"_"书写换行(有上限)
    '注意ppt创建是30帧1秒
    template_marker = template_marker + _
    Chr(10) & "<rdf:li>" _
    & Chr(10) & "<rdf:Description" _
    & Chr(10) & "xmpDM:startTime = " & Chr$(34) & list(i) * 30 & Chr$(34) _
    & Chr(10) & "xmpDM:Duration = ""0""" _
    & Chr(10) & "xmpDM:name= " & Chr$(34) & i & Chr$(34) & ">" _
    & Chr(10) & "</rdf:Description>" _
    & Chr(10) & "</rdf:li>"

Next i


'XMP
xmp_template = _
"<?xpacket begin=""?"" id=""W5M0MpCehiHzreSzNTczkc9d""?>" _
& Chr(10) & "<x:xmpmeta xmlns:x=""adobe:ns:meta/"" x:xmptk=""Adobe XMP Core 5.6-c137 79.159768, 2016/08/11-13:24:42"">" _
& Chr(10) & "<rdf:RDF xmlns:rdf=""http://www.w3.org/1999/02/22-rdf-syntax-ns#"">" _
& Chr(10) & "<rdf:Description rdf:about=""""" _
& Chr(10) & "xmlns:xmp=""http://ns.adobe.com/xap/1.0/""" _
& Chr(10) & "xmlns:xmpDM=""http://ns.adobe.com/xmp/1.0/DynamicMedia/""" _
& Chr(10) & "xmlns:xmpMM=""http://ns.adobe.com/xap/1.0/mm/""" _
& Chr(10) & "xmlns:stEvt=""http://ns.adobe.com/xap/1.0/sType/ResourceEvent#"">" _
& Chr(10) & "<xmpDM:Tracks>" _
& Chr(10) & "<rdf:Bag>" _
& Chr(10) & "<rdf:li>" _
& Chr(10) & "<rdf:Description" _
& Chr(10) & "xmpDM:trackName=""Comment""" _
& Chr(10) & "xmpDM:trackType =""Comment""" _
& Chr(10) & "xmpDM:frameRate=""f30"">" _
& Chr(10) & "<xmpDM:markers>" _
& Chr(10) & "<rdf:Seq>" _
& template_marker _
& Chr(10) & "</rdf:Seq>" _
& Chr(10) & "</xmpDM:markers>" _
& Chr(10) & "</rdf:Description>" _
& Chr(10) & "</rdf:li>" _
& Chr(10) & "</rdf:Bag>" & Chr(10) & "</xmpDM:Tracks>" & Chr(10) & "</rdf:Description>" & Chr(10) & "</rdf:RDF>" & Chr(10) & "</x:xmpmeta>" _
& Chr(10) & "<?xpacket end=""w""?>"


'文件导出
'file = VBA.Replace(ActivePresentation.FullName, ".pptm", "") & ".xmp"
file = VBA.Split(ActivePresentation.FullName, ".")(0) & ".xmp"

If Dir(file) <> "" Then
    Kill file
End If

Open file For Output As #1
    Print #1, xmp_template
    Debug.Print "Complate"
Close #1


End Sub
