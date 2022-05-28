Attribute VB_Name = "CreateXMP"
Sub CreateXMP()

Dim tmpTypa As Integer
Dim tmpDelayTime As Single
Dim tmpDuration As Single
Dim tmpRepeatTime As Single
Dim tmpMaster As Single
Dim tmpMaster_up As Single
Dim tmpTextBoxDuration As Single
Dim i As Integer
Dim j As Integer
Dim e As Single '总时长
Dim emax As Single
Dim template_marker As String
Dim xmp_template As String
Dim file As String
  
  
'Main
For i = 1 To ActivePresentation.Slides.Count

    '加入幻灯片切换时间(排除效果“无”) 每切换下一个幻灯片会有2fps误差,加 2/30fps = 0.067s
    If ActivePresentation.Slides(i).SlideShowTransition.EntryEffect <> 0 Then
        e = e + ActivePresentation.Slides(i).SlideShowTransition.Duration + 0.67
    End If
    
    '加入切换标记
    template_marker = template_marker + _
    Chr(10) & "<rdf:li>" _
    & Chr(10) & "<rdf:Description" _
    & Chr(10) & "xmpDM:startTime = " & Chr$(34) & e * 30 & Chr$(34) _
    & Chr(10) & "xmpDM:Duration = ""0""" _
    & Chr(10) & "xmpDM:name= " & Chr$(34) & i & Chr$(34) & ">" _
    & Chr(10) & "</rdf:Description>" _
    & Chr(10) & "</rdf:li>"

    
    For j = 1 To ActivePresentation.Slides(i).TimeLine.MainSequence.Count
        
        With ActivePresentation.Slides(i).TimeLine.MainSequence(j)
            tmpType = .Timing.TriggerType
            tmpDelayTime = .Timing.TriggerDelayTime
            tmpDuration = .Timing.Duration
            If .Timing.RepeatCount <> 0 Then
                tmpRepeatTime = (.Timing.RepeatCount - 1) * tmpDuration
            End If

            tmpMaster = tmpDelayTime + tmpDuration + tmpRepeatTime
            
            '加上文本逐字动画时长(文字之间延迟 50%)
            If .EffectInformation.TextUnitEffect = 1 Then
                tmpTextBoxDuration = .Timing.Duration * 0.5 * (.Shape.TextFrame.TextRange.Length - 1) '减去一个字符
                If VBA.Split(tmpTextBoxDuration, ".5")(0) <> CInt(tmpTextBoxDuration - 0.1) Then '利用该方法判断，整数和0.5的小数不加0.05
                    tmpTextBoxDuration = tmpTextBoxDuration + 0.05
                End If
                tmpMaster = tmpMaster + tmpTextBoxDuration
            End If
        End With
        
        '上一元素动画时长及类型
        If j <> 1 Then
            With ActivePresentation.Slides(i).TimeLine.MainSequence(j - 1)
                tmpType_up = .Timing.TriggerType
                tmpDelayTime_up = .Timing.TriggerDelayTime
                tmpDuration_up = .Timing.Duration
                If .Timing.RepeatCount <> 0 Then
                    tmpRepeatTime_up = (.Timing.RepeatCount - 1) * tmpDuration_up
                End If
                
                tmpMaster_up = tmpDelayTime_up + tmpDuration_up + tmpRepeatTime_up
            End With
        End If

        '类型2 上一项同时，取动画时间最长的
        If tmpType = 2 Then
            If tmpMaster > tmpMaster_up Then
                If tmpMaster > emax Then
                    emax = tmpMaster
                End If
                e = e - tmpMaster_up + emax
            End If
        End If
        
        '类型1 点击
        If tmpType = 1 Then
            e = e + tmpMaster
            '重置emax
            emax = 0
        End If
        
        '类型3 上一项后
        If tmpType = 3 Then
            e = e + tmpMaster
        End If

        Debug.Print e; emax; tmpMaster; tmpMaster_up
        
        'ASCII Chr() https://baike.baidu.com/item/Chr/580328
        '字符串末位加"_"书写换行(有上限)
        '注意ppt创建是30帧1秒
        template_marker = template_marker + _
        Chr(10) & "<rdf:li>" _
        & Chr(10) & "<rdf:Description" _
        & Chr(10) & "xmpDM:startTime = " & Chr$(34) & e * 30 & Chr$(34) _
        & Chr(10) & "xmpDM:Duration = ""0""" _
        & Chr(10) & "xmpDM:name= " & Chr$(34) & i & Chr$(34) & ">" _
        & Chr(10) & "</rdf:Description>" _
        & Chr(10) & "</rdf:li>"
        
    Next j
    
    '创建视频幻灯片每张时间(已弃，请设置为0秒）
    '因创建视频会均衡动画时常分布到5s内，导致标记不准确
    'Dim list(364) As Single
    'Dim f As Single
    'list(i) = e
    'f = list(i) - list(i - 1) '求当页动画时长
    'If f < 5 Then
        'e = e + 5 - f
        'Debug.Print "not"; f
    'End If
    
    '变更值
    'list(i) = e
    
    '提取单页时间
    Debug.Print "第"; i; "页"; e; "秒"

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
