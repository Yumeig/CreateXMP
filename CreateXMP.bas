Attribute VB_Name = "CreateXMP"
Sub CreateXMP()

Dim tmpTypa As Integer
Dim tmpDelayTime As Single
Dim tmpDuration As Single
Dim tmpRepeatTime As Single
Dim tmpMaster As Single
Dim tmpMaster_up As Single
Dim tmpTextBoxStr As String
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
    
    'ÿ�л���һ���õ�Ƭ����2fps�����,�� 2/30fps = 0.066s(������֡���йأ�����ݵ�����Ƶʱ���������Ĵ�ֵ�������)
    If i <> 1 Then
        e = e + 0.066 '30.303fps 0.133s, 62.5fps 0.066s
    End If
    
    '�����л����
    template_marker = template_marker + _
    vbLf & "<rdf:li>" _
    & vbLf & "<rdf:Description" _
    & vbLf & "xmpDM:startTime = " & Chr$(34) & e * 30 & Chr$(34) _
    & vbLf & "xmpDM:Duration = ""0""" _
    & vbLf & "xmpDM:name= " & Chr$(34) & i & Chr$(34) & ">" _
    & vbLf & "</rdf:Description>" _
    & vbLf & "</rdf:li>"

    '����dValue
    dValue = 0
    
    For j = 1 To ActivePresentation.Slides(i).TimeLine.MainSequence.Count
        
        With ActivePresentation.Slides(i).TimeLine.MainSequence(j)
            tmpType = .Timing.TriggerType
            tmpDelayTime = .Timing.TriggerDelayTime
            tmpDuration = .Timing.Duration
            If .Timing.RepeatCount <> 0 Then
                tmpRepeatTime = (.Timing.RepeatCount - 1) * tmpDuration
            End If

            tmpMaster = tmpDelayTime + tmpDuration + tmpRepeatTime
            
            '�����ı����ֶ���ʱ��(����֮���ӳ� 50%)
            If .EffectInformation.TextUnitEffect = 1 Then
                tmpTextBoxStr = VBA.Replace(VBA.Replace(VBA.Replace(.Shape.TextFrame.TextRange.Text, " ", ""), vbTab, ""), vbCr, "")
                tmpTextBoxDuration = .Timing.Duration * 0.5 * (VBA.Len(tmpTextBoxStr) - 1) '��ȥһ���ַ�����
                'If VBA.Split(tmpTextBoxDuration, ".5")(0) <> CInt(tmpTextBoxDuration - 0.1) Then '���ø÷����жϣ�������0.5��С������0.05(����������ʵ���ǻᱣ����λС�����м����)
                    'tmpTextBoxDuration = tmpTextBoxDuration + 0.05
                'End If
                tmpMaster = tmpMaster + tmpTextBoxDuration
            End If
        End With
        
        '��һԪ�ض���ʱ��������
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

        '����2 ��һ��ͬʱ��ȡ����ʱ�����
        If tmpType = 2 Then
            If tmpMaster < tmpMaster_up Then
                dValue = emax - tmpMaster '��ֵ
            End If
            
            If tmpMaster > tmpMaster_up Then
                If tmpMaster > emax Then
                    emax = tmpMaster
                    e = e - (tmpMaster_up + dValue) + emax
                End If
                '����dValue
                dValue = 0
            End If
        End If
        
        '����1 ���
        If tmpType = 1 Then
            e = e + tmpMaster
            '����emax
            emax = tmpMaster
        End If
        
        '����3 ��һ���
        If tmpType = 3 Then
            e = e + tmpMaster
            '����emax
            emax = tmpMaster
        End If

        Debug.Print e; emax; tmpMaster; tmpMaster_up
        
        'ASCII Chr() https://baike.baidu.com/item/Chr/580328
        '�ַ���ĩλ��"_"��д����(������)
        'ע��ppt������30֡1��
        template_marker = template_marker + _
        vbLf & "<rdf:li>" _
        & vbLf & "<rdf:Description" _
        & vbLf & "xmpDM:startTime = " & Chr(34) & e * 30 & Chr(34) _
        & vbLf & "xmpDM:Duration = ""0""" _
        & vbLf & "xmpDM:name= " & Chr(34) & i & Chr$(34) & ">" _
        & vbLf & "</rdf:Description>" _
        & vbLf & "</rdf:li>"
        
    Next j
    
    '������Ƶ�õ�Ƭÿ��ʱ��(������������Ϊ0�룩
    '�򴴽���Ƶ����⶯��ʱ���ֲ���5s�ڣ����±�ǲ�׼ȷ
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
& vbLf & "<x:xmpmeta xmlns:x=""adobe:ns:meta/"" x:xmptk=""Adobe XMP Core 5.6-c137 79.159768, 2016/08/11-13:24:42"">" _
& vbLf & "<rdf:RDF xmlns:rdf=""http://www.w3.org/1999/02/22-rdf-syntax-ns#"">" _
& vbLf & "<rdf:Description rdf:about=""""" _
& vbLf & "xmlns:xmp=""http://ns.adobe.com/xap/1.0/""" _
& vbLf & "xmlns:xmpDM=""http://ns.adobe.com/xmp/1.0/DynamicMedia/""" _
& vbLf & "xmlns:xmpMM=""http://ns.adobe.com/xap/1.0/mm/""" _
& vbLf & "xmlns:stEvt=""http://ns.adobe.com/xap/1.0/sType/ResourceEvent#"">" _
& vbLf & "<xmpDM:Tracks>" _
& vbLf & "<rdf:Bag>" _
& vbLf & "<rdf:li>" _
& vbLf & "<rdf:Description" _
& vbLf & "xmpDM:trackName=""Comment""" _
& vbLf & "xmpDM:trackType =""Comment""" _
& vbLf & "xmpDM:frameRate=""f30"">" _
& vbLf & "<xmpDM:markers>" _
& vbLf & "<rdf:Seq>" _
& template_marker _
& vbLf & "</rdf:Seq>" _
& vbLf & "</xmpDM:markers>" _
& vbLf & "</rdf:Description>" _
& vbLf & "</rdf:li>" _
& vbLf & "</rdf:Bag>" & vbLf & "</xmpDM:Tracks>" & vbLf & "</rdf:Description>" & vbLf & "</rdf:RDF>" & vbLf & "</x:xmpmeta>" _
& vbLf & "<?xpacket end=""w""?>"

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
