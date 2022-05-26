Sub CreateXMP()

Dim tmpDelayTime As Single
Dim tmpDuration As Single
Dim tmpRepeatTime As Single
Dim tmpMaster As Single
Dim tmpTypa As Integer
Dim tmpTextBoxDuration As Single
Dim i As Integer
Dim j As Integer
Dim e As Single '��ʱ��
Dim emax As Single
Dim template_marker As String
Dim xmp_template As String
Dim file As String


'Main
For i = 1 To ActivePresentation.Slides.Count

    '����õ�Ƭ�л�ʱ��(�ų�Ч�����ޡ�)
    If ActivePresentation.Slides(i).SlideShowTransition.EntryEffect <> 0 Then
        e = e + ActivePresentation.Slides(i).SlideShowTransition.Duration
    End If
    
    '�����л����
    template_marker = template_marker + _
    Chr(10) & "<rdf:li>" _
    & Chr(10) & "<rdf:Description" _
    & Chr(10) & "xmpDM:startTime = " & Chr$(34) & e * 30 & Chr$(34) _
    & Chr(10) & "xmpDM:Duration = ""0""" _
    & Chr(10) & "xmpDM:name= " & Chr$(34) & i & Chr$(34) & ">" _
    & Chr(10) & "</rdf:Description>" _
    & Chr(10) & "</rdf:li>"
    
    
    For j = 1 To ActivePresentation.Slides.Item(i).TimeLine.MainSequence.Count
        
        tmpType = ActivePresentation.Slides.Item(i).TimeLine.MainSequence.Item(j).Timing.TriggerType
        tmpDelayTime = ActivePresentation.Slides.Item(i).TimeLine.MainSequence.Item(j).Timing.TriggerDelayTime
        tmpDuration = ActivePresentation.Slides.Item(i).TimeLine.MainSequence.Item(j).Timing.Duration
        tmpRepeatTime = (ActivePresentation.Slides.Item(i).TimeLine.MainSequence.Item(j).Timing.RepeatCount - 1) * tmpDuration
        tmpMaster = tmpDelayTime + tmpDuration + tmpRepeatTime
        
        '�����ı����ֶ���ʱ��(����֮���ӳ� 50%) ����Ӧ��ȡ��״������Ӧ��������(δʵ�֣����������ֶ������㼶��Ӧ��������)
        If ActivePresentation.Slides(i).TimeLine.MainSequence(j).EffectInformation.TextUnitEffect = 1 Then
            tmpTextBoxDuration = ActivePresentation.Slides(i).TimeLine.MainSequence(j).Timing.Duration * 0.5 * (ActivePresentation.Slides(i).Shapes(j).TextFrame.TextRange.Length - 1) '��ȥһ���ַ�
            tmpMaster = tmpMaster + tmpTextBoxDuration
        End If

        '��һԪ�ض���ʱ��������
        If j <> 1 Then
            tmpType_up = ActivePresentation.Slides.Item(i).TimeLine.MainSequence.Item(j - 1).Timing.TriggerType
            tmpDelayTime_up = ActivePresentation.Slides.Item(i).TimeLine.MainSequence.Item(j - 1).Timing.TriggerDelayTime
            tmpDuration_up = ActivePresentation.Slides.Item(i).TimeLine.MainSequence.Item(j - 1).Timing.Duration
            tmpRepeatTime_up = (ActivePresentation.Slides.Item(i).TimeLine.MainSequence.Item(j - 1).Timing.RepeatCount - 1) * tmpDuration_up
            tmpMaster_up = tmpDelayTime_up + tmpDuration_up + tmpRepeatTime_up
        End If

        '����2 ��һ��ͬʱ��ȡ����ʱ�����
        If tmpType = 2 Then
            If tmpMaster > tmpMaster_up Then
                If tmpMaster > emax Then
                    emax = tmpMaster
                End If
                e = emax
            End If
        End If
        
        '����1 ���
        If tmpType = 1 Then
            e = e + tmpMaster
            '����emax
            emax = 0
        End If
        
        '����3 ��һ���
        If tmpType = 3 Then
            e = e + tmpMaster
        End If
        
        Debug.Print e; emax; tmpMaster; tmpMaster_up
        
        'ASCII Chr() https://baike.baidu.com/item/Chr/580328
        '�ַ���ĩλ��"_"��д����(������)
        'ע��ppt������30֡1��
        template_marker = template_marker + _
        Chr(10) & "<rdf:li>" _
        & Chr(10) & "<rdf:Description" _
        & Chr(10) & "xmpDM:startTime = " & Chr$(34) & e * 30 & Chr$(34) _
        & Chr(10) & "xmpDM:Duration = ""0""" _
        & Chr(10) & "xmpDM:name= " & Chr$(34) & i & Chr$(34) & ">" _
        & Chr(10) & "</rdf:Description>" _
        & Chr(10) & "</rdf:li>"
        
    Next j
    
    '������Ƶ�õ�Ƭÿ��ʱ��(������������Ϊ0�룩
    '�綯��ʱ��С��5s��������Ƶ����⶯��ʱ���ֲ���5s�ڣ����±�ǲ�׼ȷ
    'Dim list(364) As Single
    'Dim f As Single
    'list(i) = e
    'f = list(i) - list(i - 1) '��ҳ����ʱ��
    'If f < 5 Then
        'e = e + 5 - f
        'Debug.Print "not"; f
    'End If
    
    '���ֵ
    'list(i) = e
    
    '��ȡ��ҳʱ��
    Debug.Print "��"; i; "ҳ"; e; "��"

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

'�ļ�����
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
