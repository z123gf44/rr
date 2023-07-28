Attribute VB_Name = "Module3"
Option Explicit
Public Function CRC_keycode(getStrall As String) As String
   ' CRC16_char(p[0]); //加密计算
    CRC_keycode = Chr_CRCkeycode(getStrall)
End Function
Public Function CRC16_keycodedata(getStrall As String) As String
   ' CRC16_char(p[0]); //加密计算
     CRC16_keycodedata = Chr16_CRCkeycode(getStrall)
    CRC16_keycodedata = CStr(CRC16_keycodedata)
End Function
Public Function Chr_CRCkeycode(data As String) As String '*RTU方式的CRC校验计算
 
Dim CrcJ, yyas As Long
Dim I As Integer
Dim j As Integer
Dim mystr() As String
ReDim mystr(Len(data))
For I = 1 To (Len(data) / 2)
    mystr(I) = Mid(data, I * 2 - 1, 2)
Next
      CrcJ = 65535                                      '*CRCj赋值65535
          yyas = Val("&H" & mystr(1))
          CrcJ = CrcJ Xor yyas                          '*每个字节与CRCj的当前值异或
          For j = 0 To 7                                 '*循环八次
              If CrcJ Mod 2 = 1 Then                     '*如果CRCj当前值除2的余数等于1  判断最低位是否为1
                 CrcJ = CrcJ \ 2                         '*则CRCj当前值除2  右移
                 CrcJ = CrcJ Xor 40961                   '*CRCj当前值与40961异或
              Else
                 CrcJ = CrcJ \ 2                         '*如果CRCj当前值最低位不等于1 则只右移
              End If
         Next j
          yyas = Val("&H" & "01")
          CrcJ = CrcJ Xor yyas                          '*每个字节与CRCj的当前值异或
          For j = 0 To 7                                 '*循环八次
              If CrcJ Mod 2 = 1 Then                     '*如果CRCj当前值除2的余数等于1  判断最低位是否为1
                 CrcJ = CrcJ \ 2                         '*则CRCj当前值除2  右移
                 CrcJ = CrcJ Xor 40961                   '*CRCj当前值与40961异或
              Else
                 CrcJ = CrcJ \ 2                         '*如果CRCj当前值最低位不等于1 则只右移
              End If
         Next j
           yyas = Val("&H" & mystr(2))
          CrcJ = CrcJ Xor yyas                          '*每个字节与CRCj的当前值异或
          For j = 0 To 7                                 '*循环八次
              If CrcJ Mod 2 = 1 Then                     '*如果CRCj当前值除2的余数等于1  判断最低位是否为1
                 CrcJ = CrcJ \ 2                         '*则CRCj当前值除2  右移
                 CrcJ = CrcJ Xor 40961                   '*CRCj当前值与40961异或
              Else
                 CrcJ = CrcJ \ 2                         '*如果CRCj当前值最低位不等于1 则只右移
              End If
         Next j
          yyas = Val("&H" & "00")
          CrcJ = CrcJ Xor yyas                          '*每个字节与CRCj的当前值异或
          For j = 0 To 7                                 '*循环八次
              If CrcJ Mod 2 = 1 Then                     '*如果CRCj当前值除2的余数等于1  判断最低位是否为1
                 CrcJ = CrcJ \ 2                         '*则CRCj当前值除2  右移
                 CrcJ = CrcJ Xor 40961                   '*CRCj当前值与40961异或
              Else
                 CrcJ = CrcJ \ 2                         '*如果CRCj当前值最低位不等于1 则只右移
              End If
         Next j
         
          yyas = Val("&H" & mystr(3))
          CrcJ = CrcJ Xor yyas                          '*每个字节与CRCj的当前值异或
          For j = 0 To 7                                 '*循环八次
              If CrcJ Mod 2 = 1 Then                     '*如果CRCj当前值除2的余数等于1  判断最低位是否为1
                 CrcJ = CrcJ \ 2                         '*则CRCj当前值除2  右移
                 CrcJ = CrcJ Xor 40961                   '*CRCj当前值与40961异或
              Else
                 CrcJ = CrcJ \ 2                         '*如果CRCj当前值最低位不等于1 则只右移
              End If
         Next j
          yyas = Val("&H" & "04")
          CrcJ = CrcJ Xor yyas                          '*每个字节与CRCj的当前值异或
          For j = 0 To 7                                 '*循环八次
              If CrcJ Mod 2 = 1 Then                     '*如果CRCj当前值除2的余数等于1  判断最低位是否为1
                 CrcJ = CrcJ \ 2                         '*则CRCj当前值除2  右移
                 CrcJ = CrcJ Xor 40961                   '*CRCj当前值与40961异或
              Else
                 CrcJ = CrcJ \ 2                         '*如果CRCj当前值最低位不等于1 则只右移
              End If
         Next j
           yyas = Val("&H" & mystr(4))
          CrcJ = CrcJ Xor yyas                          '*每个字节与CRCj的当前值异或
          For j = 0 To 7                                 '*循环八次
              If CrcJ Mod 2 = 1 Then                     '*如果CRCj当前值除2的余数等于1  判断最低位是否为1
                 CrcJ = CrcJ \ 2                         '*则CRCj当前值除2  右移
                 CrcJ = CrcJ Xor 40961                   '*CRCj当前值与40961异或
              Else
                 CrcJ = CrcJ \ 2                         '*如果CRCj当前值最低位不等于1 则只右移
              End If
         Next j
          yyas = Val("&H" & "04")
          CrcJ = CrcJ Xor yyas                          '*每个字节与CRCj的当前值异或
          For j = 0 To 7                                 '*循环八次
              If CrcJ Mod 2 = 1 Then                     '*如果CRCj当前值除2的余数等于1  判断最低位是否为1
                 CrcJ = CrcJ \ 2                         '*则CRCj当前值除2  右移
                 CrcJ = CrcJ Xor 40961                   '*CRCj当前值与40961异或
              Else
                 CrcJ = CrcJ \ 2                         '*如果CRCj当前值最低位不等于1 则只右移
              End If
         Next j
         
           yyas = Val("&H" & mystr(5))
          CrcJ = CrcJ Xor yyas                          '*每个字节与CRCj的当前值异或
          For j = 0 To 7                                 '*循环八次
              If CrcJ Mod 2 = 1 Then                     '*如果CRCj当前值除2的余数等于1  判断最低位是否为1
                 CrcJ = CrcJ \ 2                         '*则CRCj当前值除2  右移
                 CrcJ = CrcJ Xor 40961                   '*CRCj当前值与40961异或
              Else
                 CrcJ = CrcJ \ 2                         '*如果CRCj当前值最低位不等于1 则只右移
              End If
         Next j
           yyas = Val("&H" & "09")
          CrcJ = CrcJ Xor yyas                          '*每个字节与CRCj的当前值异或
          For j = 0 To 7                                 '*循环八次
              If CrcJ Mod 2 = 1 Then                     '*如果CRCj当前值除2的余数等于1  判断最低位是否为1
                 CrcJ = CrcJ \ 2                         '*则CRCj当前值除2  右移
                 CrcJ = CrcJ Xor 40961                   '*CRCj当前值与40961异或
              Else
                 CrcJ = CrcJ \ 2                         '*如果CRCj当前值最低位不等于1 则只右移
              End If
         Next j
          yyas = Val("&H" & "04")
          CrcJ = CrcJ Xor yyas                          '*每个字节与CRCj的当前值异或
          For j = 0 To 7                                 '*循环八次
              If CrcJ Mod 2 = 1 Then                     '*如果CRCj当前值除2的余数等于1  判断最低位是否为1
                 CrcJ = CrcJ \ 2                         '*则CRCj当前值除2  右移
                 CrcJ = CrcJ Xor 40961                   '*CRCj当前值与40961异或
              Else
                 CrcJ = CrcJ \ 2                         '*如果CRCj当前值最低位不等于1 则只右移
              End If
         Next j
         
           yyas = Val("&H" & mystr(6))
          CrcJ = CrcJ Xor yyas                          '*每个字节与CRCj的当前值异或
          For j = 0 To 7                                 '*循环八次
              If CrcJ Mod 2 = 1 Then                     '*如果CRCj当前值除2的余数等于1  判断最低位是否为1
                 CrcJ = CrcJ \ 2                         '*则CRCj当前值除2  右移
                 CrcJ = CrcJ Xor 40961                   '*CRCj当前值与40961异或
              Else
                 CrcJ = CrcJ \ 2                         '*如果CRCj当前值最低位不等于1 则只右移
              End If
         Next j
          yyas = Val("&H" & mystr(7))
          CrcJ = CrcJ Xor yyas                          '*每个字节与CRCj的当前值异或
          For j = 0 To 7                                 '*循环八次
              If CrcJ Mod 2 = 1 Then                     '*如果CRCj当前值除2的余数等于1  判断最低位是否为1
                 CrcJ = CrcJ \ 2                         '*则CRCj当前值除2  右移
                 CrcJ = CrcJ Xor 40961                   '*CRCj当前值与40961异或
              Else
                 CrcJ = CrcJ \ 2                         '*如果CRCj当前值最低位不等于1 则只右移
              End If
         Next j
 
           yyas = Val("&H" & mystr(8))
          CrcJ = CrcJ Xor yyas                          '*每个字节与CRCj的当前值异或
          For j = 0 To 7                                 '*循环八次
              If CrcJ Mod 2 = 1 Then                     '*如果CRCj当前值除2的余数等于1  判断最低位是否为1
                 CrcJ = CrcJ \ 2                         '*则CRCj当前值除2  右移
                 CrcJ = CrcJ Xor 40961                   '*CRCj当前值与40961异或
              Else
                 CrcJ = CrcJ \ 2                         '*如果CRCj当前值最低位不等于1 则只右移
              End If
         Next j
Chr_CRCkeycode = ""
Chr_CRCkeycode = Right("0000" + Hex(CrcJ), 4)             '*计算的CRCj当前值转换为十六进制
Chr_CRCkeycode = Right(Chr_CRCkeycode, 2) + Left(Chr_CRCkeycode, 2)     '*低字节放在前面，高字节放在后面
     
        CrcJ = 65535                                       '*CRCj赋值65535
          yyas = Val("&H" & mystr(1))
          CrcJ = CrcJ Xor yyas                          '*每个字节与CRCj的当前值异或
          For j = 0 To 7                                 '*循环八次
              If CrcJ Mod 2 = 1 Then                     '*如果CRCj当前值除2的余数等于1  判断最低位是否为1
                 CrcJ = CrcJ \ 2                         '*则CRCj当前值除2  右移
                 CrcJ = CrcJ Xor 40961                   '*CRCj当前值与40961异或
              Else
                 CrcJ = CrcJ \ 2                         '*如果CRCj当前值最低位不等于1 则只右移
              End If
         Next j
          yyas = Val("&H" & "00")
          CrcJ = CrcJ Xor yyas                          '*每个字节与CRCj的当前值异或
          For j = 0 To 7                                 '*循环八次
              If CrcJ Mod 2 = 1 Then                     '*如果CRCj当前值除2的余数等于1  判断最低位是否为1
                 CrcJ = CrcJ \ 2                         '*则CRCj当前值除2  右移
                 CrcJ = CrcJ Xor 40961                   '*CRCj当前值与40961异或
              Else
                 CrcJ = CrcJ \ 2                         '*如果CRCj当前值最低位不等于1 则只右移
              End If
         Next j
           yyas = Val("&H" & mystr(2))
          CrcJ = CrcJ Xor yyas                          '*每个字节与CRCj的当前值异或
          For j = 0 To 7                                 '*循环八次
              If CrcJ Mod 2 = 1 Then                     '*如果CRCj当前值除2的余数等于1  判断最低位是否为1
                 CrcJ = CrcJ \ 2                         '*则CRCj当前值除2  右移
                 CrcJ = CrcJ Xor 40961                   '*CRCj当前值与40961异或
              Else
                 CrcJ = CrcJ \ 2                         '*如果CRCj当前值最低位不等于1 则只右移
              End If
         Next j
          yyas = Val("&H" & "07")
          CrcJ = CrcJ Xor yyas                          '*每个字节与CRCj的当前值异或
          For j = 0 To 7                                 '*循环八次
              If CrcJ Mod 2 = 1 Then                     '*如果CRCj当前值除2的余数等于1  判断最低位是否为1
                 CrcJ = CrcJ \ 2                         '*则CRCj当前值除2  右移
                 CrcJ = CrcJ Xor 40961                   '*CRCj当前值与40961异或
              Else
                 CrcJ = CrcJ \ 2                         '*如果CRCj当前值最低位不等于1 则只右移
              End If
         Next j
         
          yyas = Val("&H" & mystr(3))
          CrcJ = CrcJ Xor yyas                          '*每个字节与CRCj的当前值异或
          For j = 0 To 7                                 '*循环八次
              If CrcJ Mod 2 = 1 Then                     '*如果CRCj当前值除2的余数等于1  判断最低位是否为1
                 CrcJ = CrcJ \ 2                         '*则CRCj当前值除2  右移
                 CrcJ = CrcJ Xor 40961                   '*CRCj当前值与40961异或
              Else
                 CrcJ = CrcJ \ 2                         '*如果CRCj当前值最低位不等于1 则只右移
              End If
         Next j
          yyas = Val("&H" & "04")
          CrcJ = CrcJ Xor yyas                          '*每个字节与CRCj的当前值异或
          For j = 0 To 7                                 '*循环八次
              If CrcJ Mod 2 = 1 Then                     '*如果CRCj当前值除2的余数等于1  判断最低位是否为1
                 CrcJ = CrcJ \ 2                         '*则CRCj当前值除2  右移
                 CrcJ = CrcJ Xor 40961                   '*CRCj当前值与40961异或
              Else
                 CrcJ = CrcJ \ 2                         '*如果CRCj当前值最低位不等于1 则只右移
              End If
         Next j
           yyas = Val("&H" & mystr(4))
          CrcJ = CrcJ Xor yyas                          '*每个字节与CRCj的当前值异或
          For j = 0 To 7                                 '*循环八次
              If CrcJ Mod 2 = 1 Then                     '*如果CRCj当前值除2的余数等于1  判断最低位是否为1
                 CrcJ = CrcJ \ 2                         '*则CRCj当前值除2  右移
                 CrcJ = CrcJ Xor 40961                   '*CRCj当前值与40961异或
              Else
                 CrcJ = CrcJ \ 2                         '*如果CRCj当前值最低位不等于1 则只右移
              End If
         Next j
          yyas = Val("&H" & "04")
          CrcJ = CrcJ Xor yyas                          '*每个字节与CRCj的当前值异或
          For j = 0 To 7                                 '*循环八次
              If CrcJ Mod 2 = 1 Then                     '*如果CRCj当前值除2的余数等于1  判断最低位是否为1
                 CrcJ = CrcJ \ 2                         '*则CRCj当前值除2  右移
                 CrcJ = CrcJ Xor 40961                   '*CRCj当前值与40961异或
              Else
                 CrcJ = CrcJ \ 2                         '*如果CRCj当前值最低位不等于1 则只右移
              End If
         Next j
         
           yyas = Val("&H" & mystr(5))
          CrcJ = CrcJ Xor yyas                          '*每个字节与CRCj的当前值异或
          For j = 0 To 7                                 '*循环八次
              If CrcJ Mod 2 = 1 Then                     '*如果CRCj当前值除2的余数等于1  判断最低位是否为1
                 CrcJ = CrcJ \ 2                         '*则CRCj当前值除2  右移
                 CrcJ = CrcJ Xor 40961                   '*CRCj当前值与40961异或
              Else
                 CrcJ = CrcJ \ 2                         '*如果CRCj当前值最低位不等于1 则只右移
              End If
         Next j
           yyas = Val("&H" & "09")
          CrcJ = CrcJ Xor yyas                          '*每个字节与CRCj的当前值异或
          For j = 0 To 7                                 '*循环八次
              If CrcJ Mod 2 = 1 Then                     '*如果CRCj当前值除2的余数等于1  判断最低位是否为1
                 CrcJ = CrcJ \ 2                         '*则CRCj当前值除2  右移
                 CrcJ = CrcJ Xor 40961                   '*CRCj当前值与40961异或
              Else
                 CrcJ = CrcJ \ 2                         '*如果CRCj当前值最低位不等于1 则只右移
              End If
         Next j
          yyas = Val("&H" & "04")
          CrcJ = CrcJ Xor yyas                          '*每个字节与CRCj的当前值异或
          For j = 0 To 7                                 '*循环八次
              If CrcJ Mod 2 = 1 Then                     '*如果CRCj当前值除2的余数等于1  判断最低位是否为1
                 CrcJ = CrcJ \ 2                         '*则CRCj当前值除2  右移
                 CrcJ = CrcJ Xor 40961                   '*CRCj当前值与40961异或
              Else
                 CrcJ = CrcJ \ 2                         '*如果CRCj当前值最低位不等于1 则只右移
              End If
         Next j
         
        
Dim strtemp As String
strtemp = Right("0000" + Hex(CrcJ), 4)             '*计算的CRCj当前值转换为十六进制
strtemp = Right(strtemp, 2) + Left(strtemp, 2)     '*低字节放在前面，高字节放在后面
Chr_CRCkeycode = Chr_CRCkeycode & strtemp
     
End Function
Public Function Chr16_CRCkeycode(data As String) As String '*RTU方式的CRC校验计算
 
Dim CrcJ, yyas As Long
Dim I As Integer
Dim j As Integer
Dim mystr() As String
ReDim mystr(Len(data))
CrcJ = 12254
For I = 1 To ((Len(data) - 1) / 2)
     mystr(I) = Mid(data, I * 2, 2)  '*CRCj赋值65535
     yyas = Val("&H" & mystr(I))
     CrcJ = CrcJ Xor yyas                          '*每个字节与CRCj的当前值异或
     For j = 0 To 7                                 '*循环八次
         If CrcJ Mod 2 = 1 Then                     '*如果CRCj当前值除2的余数等于1  判断最低位是否为1
            CrcJ = CrcJ \ 2                         '*则CRCj当前值除2  右移
            CrcJ = CrcJ Xor 40961                   '*CRCj当前值与40961异或
         Else
            CrcJ = CrcJ \ 2                         '*如果CRCj当前值最低位不等于1 则只右移
         End If
    Next j
Next
     yyas = Val("&H" & "D0")
     CrcJ = CrcJ Xor yyas                          '*每个字节与CRCj的当前值异或
     For j = 0 To 7                                 '*循环八次
         If CrcJ Mod 2 = 1 Then                     '*如果CRCj当前值除2的余数等于1  判断最低位是否为1
            CrcJ = CrcJ \ 2                         '*则CRCj当前值除2  右移
            CrcJ = CrcJ Xor 40961                   '*CRCj当前值与40961异或
         Else
            CrcJ = CrcJ \ 2                         '*如果CRCj当前值最低位不等于1 则只右移
         End If
    Next j
    
Chr16_CRCkeycode = ""
Dim strtemp As String
strtemp = Right("0000" + Hex(CrcJ), 4)             '*计算的CRCj当前值转换为十六进制
'strtemp = Right(strtemp, 2) + Left(strtemp, 2)     '*低字节放在前面，高字节放在后面
Chr16_CRCkeycode = Chr16_CRCkeycode & strtemp
End Function

Public Function My_msgbox(data As String)   '*RTU方式的CRC校验计算
    If jingdu1 = 0 Then
     MsgBox (data)
    End If
End Function
Public Function clean_disbox()    ' 清除 弹出窗
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
