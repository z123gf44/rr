Attribute VB_Name = "Module3"
Option Explicit
Public Function CRC_keycode(getStrall As String) As String
   ' CRC16_char(p[0]); //���ܼ���
    CRC_keycode = Chr_CRCkeycode(getStrall)
End Function
Public Function CRC16_keycodedata(getStrall As String) As String
   ' CRC16_char(p[0]); //���ܼ���
     CRC16_keycodedata = Chr16_CRCkeycode(getStrall)
    CRC16_keycodedata = CStr(CRC16_keycodedata)
End Function
Public Function Chr_CRCkeycode(data As String) As String '*RTU��ʽ��CRCУ�����
 
Dim CrcJ, yyas As Long
Dim I As Integer
Dim j As Integer
Dim mystr() As String
ReDim mystr(Len(data))
For I = 1 To (Len(data) / 2)
    mystr(I) = Mid(data, I * 2 - 1, 2)
Next
      CrcJ = 65535                                      '*CRCj��ֵ65535
          yyas = Val("&H" & mystr(1))
          CrcJ = CrcJ Xor yyas                          '*ÿ���ֽ���CRCj�ĵ�ǰֵ���
          For j = 0 To 7                                 '*ѭ���˴�
              If CrcJ Mod 2 = 1 Then                     '*���CRCj��ǰֵ��2����������1  �ж����λ�Ƿ�Ϊ1
                 CrcJ = CrcJ \ 2                         '*��CRCj��ǰֵ��2  ����
                 CrcJ = CrcJ Xor 40961                   '*CRCj��ǰֵ��40961���
              Else
                 CrcJ = CrcJ \ 2                         '*���CRCj��ǰֵ���λ������1 ��ֻ����
              End If
         Next j
          yyas = Val("&H" & "01")
          CrcJ = CrcJ Xor yyas                          '*ÿ���ֽ���CRCj�ĵ�ǰֵ���
          For j = 0 To 7                                 '*ѭ���˴�
              If CrcJ Mod 2 = 1 Then                     '*���CRCj��ǰֵ��2����������1  �ж����λ�Ƿ�Ϊ1
                 CrcJ = CrcJ \ 2                         '*��CRCj��ǰֵ��2  ����
                 CrcJ = CrcJ Xor 40961                   '*CRCj��ǰֵ��40961���
              Else
                 CrcJ = CrcJ \ 2                         '*���CRCj��ǰֵ���λ������1 ��ֻ����
              End If
         Next j
           yyas = Val("&H" & mystr(2))
          CrcJ = CrcJ Xor yyas                          '*ÿ���ֽ���CRCj�ĵ�ǰֵ���
          For j = 0 To 7                                 '*ѭ���˴�
              If CrcJ Mod 2 = 1 Then                     '*���CRCj��ǰֵ��2����������1  �ж����λ�Ƿ�Ϊ1
                 CrcJ = CrcJ \ 2                         '*��CRCj��ǰֵ��2  ����
                 CrcJ = CrcJ Xor 40961                   '*CRCj��ǰֵ��40961���
              Else
                 CrcJ = CrcJ \ 2                         '*���CRCj��ǰֵ���λ������1 ��ֻ����
              End If
         Next j
          yyas = Val("&H" & "00")
          CrcJ = CrcJ Xor yyas                          '*ÿ���ֽ���CRCj�ĵ�ǰֵ���
          For j = 0 To 7                                 '*ѭ���˴�
              If CrcJ Mod 2 = 1 Then                     '*���CRCj��ǰֵ��2����������1  �ж����λ�Ƿ�Ϊ1
                 CrcJ = CrcJ \ 2                         '*��CRCj��ǰֵ��2  ����
                 CrcJ = CrcJ Xor 40961                   '*CRCj��ǰֵ��40961���
              Else
                 CrcJ = CrcJ \ 2                         '*���CRCj��ǰֵ���λ������1 ��ֻ����
              End If
         Next j
         
          yyas = Val("&H" & mystr(3))
          CrcJ = CrcJ Xor yyas                          '*ÿ���ֽ���CRCj�ĵ�ǰֵ���
          For j = 0 To 7                                 '*ѭ���˴�
              If CrcJ Mod 2 = 1 Then                     '*���CRCj��ǰֵ��2����������1  �ж����λ�Ƿ�Ϊ1
                 CrcJ = CrcJ \ 2                         '*��CRCj��ǰֵ��2  ����
                 CrcJ = CrcJ Xor 40961                   '*CRCj��ǰֵ��40961���
              Else
                 CrcJ = CrcJ \ 2                         '*���CRCj��ǰֵ���λ������1 ��ֻ����
              End If
         Next j
          yyas = Val("&H" & "04")
          CrcJ = CrcJ Xor yyas                          '*ÿ���ֽ���CRCj�ĵ�ǰֵ���
          For j = 0 To 7                                 '*ѭ���˴�
              If CrcJ Mod 2 = 1 Then                     '*���CRCj��ǰֵ��2����������1  �ж����λ�Ƿ�Ϊ1
                 CrcJ = CrcJ \ 2                         '*��CRCj��ǰֵ��2  ����
                 CrcJ = CrcJ Xor 40961                   '*CRCj��ǰֵ��40961���
              Else
                 CrcJ = CrcJ \ 2                         '*���CRCj��ǰֵ���λ������1 ��ֻ����
              End If
         Next j
           yyas = Val("&H" & mystr(4))
          CrcJ = CrcJ Xor yyas                          '*ÿ���ֽ���CRCj�ĵ�ǰֵ���
          For j = 0 To 7                                 '*ѭ���˴�
              If CrcJ Mod 2 = 1 Then                     '*���CRCj��ǰֵ��2����������1  �ж����λ�Ƿ�Ϊ1
                 CrcJ = CrcJ \ 2                         '*��CRCj��ǰֵ��2  ����
                 CrcJ = CrcJ Xor 40961                   '*CRCj��ǰֵ��40961���
              Else
                 CrcJ = CrcJ \ 2                         '*���CRCj��ǰֵ���λ������1 ��ֻ����
              End If
         Next j
          yyas = Val("&H" & "04")
          CrcJ = CrcJ Xor yyas                          '*ÿ���ֽ���CRCj�ĵ�ǰֵ���
          For j = 0 To 7                                 '*ѭ���˴�
              If CrcJ Mod 2 = 1 Then                     '*���CRCj��ǰֵ��2����������1  �ж����λ�Ƿ�Ϊ1
                 CrcJ = CrcJ \ 2                         '*��CRCj��ǰֵ��2  ����
                 CrcJ = CrcJ Xor 40961                   '*CRCj��ǰֵ��40961���
              Else
                 CrcJ = CrcJ \ 2                         '*���CRCj��ǰֵ���λ������1 ��ֻ����
              End If
         Next j
         
           yyas = Val("&H" & mystr(5))
          CrcJ = CrcJ Xor yyas                          '*ÿ���ֽ���CRCj�ĵ�ǰֵ���
          For j = 0 To 7                                 '*ѭ���˴�
              If CrcJ Mod 2 = 1 Then                     '*���CRCj��ǰֵ��2����������1  �ж����λ�Ƿ�Ϊ1
                 CrcJ = CrcJ \ 2                         '*��CRCj��ǰֵ��2  ����
                 CrcJ = CrcJ Xor 40961                   '*CRCj��ǰֵ��40961���
              Else
                 CrcJ = CrcJ \ 2                         '*���CRCj��ǰֵ���λ������1 ��ֻ����
              End If
         Next j
           yyas = Val("&H" & "09")
          CrcJ = CrcJ Xor yyas                          '*ÿ���ֽ���CRCj�ĵ�ǰֵ���
          For j = 0 To 7                                 '*ѭ���˴�
              If CrcJ Mod 2 = 1 Then                     '*���CRCj��ǰֵ��2����������1  �ж����λ�Ƿ�Ϊ1
                 CrcJ = CrcJ \ 2                         '*��CRCj��ǰֵ��2  ����
                 CrcJ = CrcJ Xor 40961                   '*CRCj��ǰֵ��40961���
              Else
                 CrcJ = CrcJ \ 2                         '*���CRCj��ǰֵ���λ������1 ��ֻ����
              End If
         Next j
          yyas = Val("&H" & "04")
          CrcJ = CrcJ Xor yyas                          '*ÿ���ֽ���CRCj�ĵ�ǰֵ���
          For j = 0 To 7                                 '*ѭ���˴�
              If CrcJ Mod 2 = 1 Then                     '*���CRCj��ǰֵ��2����������1  �ж����λ�Ƿ�Ϊ1
                 CrcJ = CrcJ \ 2                         '*��CRCj��ǰֵ��2  ����
                 CrcJ = CrcJ Xor 40961                   '*CRCj��ǰֵ��40961���
              Else
                 CrcJ = CrcJ \ 2                         '*���CRCj��ǰֵ���λ������1 ��ֻ����
              End If
         Next j
         
           yyas = Val("&H" & mystr(6))
          CrcJ = CrcJ Xor yyas                          '*ÿ���ֽ���CRCj�ĵ�ǰֵ���
          For j = 0 To 7                                 '*ѭ���˴�
              If CrcJ Mod 2 = 1 Then                     '*���CRCj��ǰֵ��2����������1  �ж����λ�Ƿ�Ϊ1
                 CrcJ = CrcJ \ 2                         '*��CRCj��ǰֵ��2  ����
                 CrcJ = CrcJ Xor 40961                   '*CRCj��ǰֵ��40961���
              Else
                 CrcJ = CrcJ \ 2                         '*���CRCj��ǰֵ���λ������1 ��ֻ����
              End If
         Next j
          yyas = Val("&H" & mystr(7))
          CrcJ = CrcJ Xor yyas                          '*ÿ���ֽ���CRCj�ĵ�ǰֵ���
          For j = 0 To 7                                 '*ѭ���˴�
              If CrcJ Mod 2 = 1 Then                     '*���CRCj��ǰֵ��2����������1  �ж����λ�Ƿ�Ϊ1
                 CrcJ = CrcJ \ 2                         '*��CRCj��ǰֵ��2  ����
                 CrcJ = CrcJ Xor 40961                   '*CRCj��ǰֵ��40961���
              Else
                 CrcJ = CrcJ \ 2                         '*���CRCj��ǰֵ���λ������1 ��ֻ����
              End If
         Next j
 
           yyas = Val("&H" & mystr(8))
          CrcJ = CrcJ Xor yyas                          '*ÿ���ֽ���CRCj�ĵ�ǰֵ���
          For j = 0 To 7                                 '*ѭ���˴�
              If CrcJ Mod 2 = 1 Then                     '*���CRCj��ǰֵ��2����������1  �ж����λ�Ƿ�Ϊ1
                 CrcJ = CrcJ \ 2                         '*��CRCj��ǰֵ��2  ����
                 CrcJ = CrcJ Xor 40961                   '*CRCj��ǰֵ��40961���
              Else
                 CrcJ = CrcJ \ 2                         '*���CRCj��ǰֵ���λ������1 ��ֻ����
              End If
         Next j
Chr_CRCkeycode = ""
Chr_CRCkeycode = Right("0000" + Hex(CrcJ), 4)             '*�����CRCj��ǰֵת��Ϊʮ������
Chr_CRCkeycode = Right(Chr_CRCkeycode, 2) + Left(Chr_CRCkeycode, 2)     '*���ֽڷ���ǰ�棬���ֽڷ��ں���
     
        CrcJ = 65535                                       '*CRCj��ֵ65535
          yyas = Val("&H" & mystr(1))
          CrcJ = CrcJ Xor yyas                          '*ÿ���ֽ���CRCj�ĵ�ǰֵ���
          For j = 0 To 7                                 '*ѭ���˴�
              If CrcJ Mod 2 = 1 Then                     '*���CRCj��ǰֵ��2����������1  �ж����λ�Ƿ�Ϊ1
                 CrcJ = CrcJ \ 2                         '*��CRCj��ǰֵ��2  ����
                 CrcJ = CrcJ Xor 40961                   '*CRCj��ǰֵ��40961���
              Else
                 CrcJ = CrcJ \ 2                         '*���CRCj��ǰֵ���λ������1 ��ֻ����
              End If
         Next j
          yyas = Val("&H" & "00")
          CrcJ = CrcJ Xor yyas                          '*ÿ���ֽ���CRCj�ĵ�ǰֵ���
          For j = 0 To 7                                 '*ѭ���˴�
              If CrcJ Mod 2 = 1 Then                     '*���CRCj��ǰֵ��2����������1  �ж����λ�Ƿ�Ϊ1
                 CrcJ = CrcJ \ 2                         '*��CRCj��ǰֵ��2  ����
                 CrcJ = CrcJ Xor 40961                   '*CRCj��ǰֵ��40961���
              Else
                 CrcJ = CrcJ \ 2                         '*���CRCj��ǰֵ���λ������1 ��ֻ����
              End If
         Next j
           yyas = Val("&H" & mystr(2))
          CrcJ = CrcJ Xor yyas                          '*ÿ���ֽ���CRCj�ĵ�ǰֵ���
          For j = 0 To 7                                 '*ѭ���˴�
              If CrcJ Mod 2 = 1 Then                     '*���CRCj��ǰֵ��2����������1  �ж����λ�Ƿ�Ϊ1
                 CrcJ = CrcJ \ 2                         '*��CRCj��ǰֵ��2  ����
                 CrcJ = CrcJ Xor 40961                   '*CRCj��ǰֵ��40961���
              Else
                 CrcJ = CrcJ \ 2                         '*���CRCj��ǰֵ���λ������1 ��ֻ����
              End If
         Next j
          yyas = Val("&H" & "07")
          CrcJ = CrcJ Xor yyas                          '*ÿ���ֽ���CRCj�ĵ�ǰֵ���
          For j = 0 To 7                                 '*ѭ���˴�
              If CrcJ Mod 2 = 1 Then                     '*���CRCj��ǰֵ��2����������1  �ж����λ�Ƿ�Ϊ1
                 CrcJ = CrcJ \ 2                         '*��CRCj��ǰֵ��2  ����
                 CrcJ = CrcJ Xor 40961                   '*CRCj��ǰֵ��40961���
              Else
                 CrcJ = CrcJ \ 2                         '*���CRCj��ǰֵ���λ������1 ��ֻ����
              End If
         Next j
         
          yyas = Val("&H" & mystr(3))
          CrcJ = CrcJ Xor yyas                          '*ÿ���ֽ���CRCj�ĵ�ǰֵ���
          For j = 0 To 7                                 '*ѭ���˴�
              If CrcJ Mod 2 = 1 Then                     '*���CRCj��ǰֵ��2����������1  �ж����λ�Ƿ�Ϊ1
                 CrcJ = CrcJ \ 2                         '*��CRCj��ǰֵ��2  ����
                 CrcJ = CrcJ Xor 40961                   '*CRCj��ǰֵ��40961���
              Else
                 CrcJ = CrcJ \ 2                         '*���CRCj��ǰֵ���λ������1 ��ֻ����
              End If
         Next j
          yyas = Val("&H" & "04")
          CrcJ = CrcJ Xor yyas                          '*ÿ���ֽ���CRCj�ĵ�ǰֵ���
          For j = 0 To 7                                 '*ѭ���˴�
              If CrcJ Mod 2 = 1 Then                     '*���CRCj��ǰֵ��2����������1  �ж����λ�Ƿ�Ϊ1
                 CrcJ = CrcJ \ 2                         '*��CRCj��ǰֵ��2  ����
                 CrcJ = CrcJ Xor 40961                   '*CRCj��ǰֵ��40961���
              Else
                 CrcJ = CrcJ \ 2                         '*���CRCj��ǰֵ���λ������1 ��ֻ����
              End If
         Next j
           yyas = Val("&H" & mystr(4))
          CrcJ = CrcJ Xor yyas                          '*ÿ���ֽ���CRCj�ĵ�ǰֵ���
          For j = 0 To 7                                 '*ѭ���˴�
              If CrcJ Mod 2 = 1 Then                     '*���CRCj��ǰֵ��2����������1  �ж����λ�Ƿ�Ϊ1
                 CrcJ = CrcJ \ 2                         '*��CRCj��ǰֵ��2  ����
                 CrcJ = CrcJ Xor 40961                   '*CRCj��ǰֵ��40961���
              Else
                 CrcJ = CrcJ \ 2                         '*���CRCj��ǰֵ���λ������1 ��ֻ����
              End If
         Next j
          yyas = Val("&H" & "04")
          CrcJ = CrcJ Xor yyas                          '*ÿ���ֽ���CRCj�ĵ�ǰֵ���
          For j = 0 To 7                                 '*ѭ���˴�
              If CrcJ Mod 2 = 1 Then                     '*���CRCj��ǰֵ��2����������1  �ж����λ�Ƿ�Ϊ1
                 CrcJ = CrcJ \ 2                         '*��CRCj��ǰֵ��2  ����
                 CrcJ = CrcJ Xor 40961                   '*CRCj��ǰֵ��40961���
              Else
                 CrcJ = CrcJ \ 2                         '*���CRCj��ǰֵ���λ������1 ��ֻ����
              End If
         Next j
         
           yyas = Val("&H" & mystr(5))
          CrcJ = CrcJ Xor yyas                          '*ÿ���ֽ���CRCj�ĵ�ǰֵ���
          For j = 0 To 7                                 '*ѭ���˴�
              If CrcJ Mod 2 = 1 Then                     '*���CRCj��ǰֵ��2����������1  �ж����λ�Ƿ�Ϊ1
                 CrcJ = CrcJ \ 2                         '*��CRCj��ǰֵ��2  ����
                 CrcJ = CrcJ Xor 40961                   '*CRCj��ǰֵ��40961���
              Else
                 CrcJ = CrcJ \ 2                         '*���CRCj��ǰֵ���λ������1 ��ֻ����
              End If
         Next j
           yyas = Val("&H" & "09")
          CrcJ = CrcJ Xor yyas                          '*ÿ���ֽ���CRCj�ĵ�ǰֵ���
          For j = 0 To 7                                 '*ѭ���˴�
              If CrcJ Mod 2 = 1 Then                     '*���CRCj��ǰֵ��2����������1  �ж����λ�Ƿ�Ϊ1
                 CrcJ = CrcJ \ 2                         '*��CRCj��ǰֵ��2  ����
                 CrcJ = CrcJ Xor 40961                   '*CRCj��ǰֵ��40961���
              Else
                 CrcJ = CrcJ \ 2                         '*���CRCj��ǰֵ���λ������1 ��ֻ����
              End If
         Next j
          yyas = Val("&H" & "04")
          CrcJ = CrcJ Xor yyas                          '*ÿ���ֽ���CRCj�ĵ�ǰֵ���
          For j = 0 To 7                                 '*ѭ���˴�
              If CrcJ Mod 2 = 1 Then                     '*���CRCj��ǰֵ��2����������1  �ж����λ�Ƿ�Ϊ1
                 CrcJ = CrcJ \ 2                         '*��CRCj��ǰֵ��2  ����
                 CrcJ = CrcJ Xor 40961                   '*CRCj��ǰֵ��40961���
              Else
                 CrcJ = CrcJ \ 2                         '*���CRCj��ǰֵ���λ������1 ��ֻ����
              End If
         Next j
         
        
Dim strtemp As String
strtemp = Right("0000" + Hex(CrcJ), 4)             '*�����CRCj��ǰֵת��Ϊʮ������
strtemp = Right(strtemp, 2) + Left(strtemp, 2)     '*���ֽڷ���ǰ�棬���ֽڷ��ں���
Chr_CRCkeycode = Chr_CRCkeycode & strtemp
     
End Function
Public Function Chr16_CRCkeycode(data As String) As String '*RTU��ʽ��CRCУ�����
 
Dim CrcJ, yyas As Long
Dim I As Integer
Dim j As Integer
Dim mystr() As String
ReDim mystr(Len(data))
CrcJ = 12254
For I = 1 To ((Len(data) - 1) / 2)
     mystr(I) = Mid(data, I * 2, 2)  '*CRCj��ֵ65535
     yyas = Val("&H" & mystr(I))
     CrcJ = CrcJ Xor yyas                          '*ÿ���ֽ���CRCj�ĵ�ǰֵ���
     For j = 0 To 7                                 '*ѭ���˴�
         If CrcJ Mod 2 = 1 Then                     '*���CRCj��ǰֵ��2����������1  �ж����λ�Ƿ�Ϊ1
            CrcJ = CrcJ \ 2                         '*��CRCj��ǰֵ��2  ����
            CrcJ = CrcJ Xor 40961                   '*CRCj��ǰֵ��40961���
         Else
            CrcJ = CrcJ \ 2                         '*���CRCj��ǰֵ���λ������1 ��ֻ����
         End If
    Next j
Next
     yyas = Val("&H" & "D0")
     CrcJ = CrcJ Xor yyas                          '*ÿ���ֽ���CRCj�ĵ�ǰֵ���
     For j = 0 To 7                                 '*ѭ���˴�
         If CrcJ Mod 2 = 1 Then                     '*���CRCj��ǰֵ��2����������1  �ж����λ�Ƿ�Ϊ1
            CrcJ = CrcJ \ 2                         '*��CRCj��ǰֵ��2  ����
            CrcJ = CrcJ Xor 40961                   '*CRCj��ǰֵ��40961���
         Else
            CrcJ = CrcJ \ 2                         '*���CRCj��ǰֵ���λ������1 ��ֻ����
         End If
    Next j
    
Chr16_CRCkeycode = ""
Dim strtemp As String
strtemp = Right("0000" + Hex(CrcJ), 4)             '*�����CRCj��ǰֵת��Ϊʮ������
'strtemp = Right(strtemp, 2) + Left(strtemp, 2)     '*���ֽڷ���ǰ�棬���ֽڷ��ں���
Chr16_CRCkeycode = Chr16_CRCkeycode & strtemp
End Function

Public Function My_msgbox(data As String)   '*RTU��ʽ��CRCУ�����
    If jingdu1 = 0 Then
     MsgBox (data)
    End If
End Function
Public Function clean_disbox()    ' ��� ������
  Delay_dis_Readsysconfig = 0
  Delay_dis_Writesysconfig = 0
  Delay_dis_ReadRegAfe = 0
  Delay_dis_WriteRegAfe = 0
  Delay_dis_Readsys2config = 0
  Delay_dis_Writesys2config = 0
  Delay_dis_EraseBalckUp = 0
  Delay_dis_Enter_Sleep_Mode = 0
  Delay_dis_Enter_WORK_Mode = 0
  Delay_dis_SetFET = 0
  Delay_dis_ReadSOC_OCV = 0
  Delay_dis_WriteSOC_OCV = 0
  Delay_dis_Readcap = 0
  Delay_dis_Writecap = 0
  Delay_dis_CALIB_RTC = 0
  Delay_dis_CALIB_VOLTAGE = 0
  Delay_dis_CALIB_CURRENT = 0
  Delay_dis_CALIB_Temp = 0
  Delay_dis_ReadBalckUp = 0
  Delay_dis_ReadMcuRAM = 0
  Delay_waite_muc_back_cmd = 0
End Function
