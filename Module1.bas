Attribute VB_Name = "Module1"
Public com_in_buf(21) As Byte, data_to_send(200) As Byte, yifawan As Integer, state As Integer
Public success As Integer, ttrr As Integer
Public s(1) As Byte, ReturnData(1) As Byte, datawait As Integer, jieshouzi As Integer

Public Getdis As String
Public SentPussFlag As Byte  '=1 暂停 =0发送
 
Public GetheairF As Boolean ' 接收到 数据头
 
Public tx_cnt As Integer
Public rx_cnt As Integer
Public rxStrall As String

  Public CRC16Lo As Byte, CRC16Hi As Byte      'CRC寄存器
Function CRC_oneByte(data As Byte)
      Dim CL As Byte, CH As Byte                '多项式码&HA001
      Dim SaveHi As Byte, SaveLo As Byte
      Dim I As Integer
      Dim Flag As Integer
 
      CL = &H1
      CH = &HA0

        CRC16Lo = CRC16Lo Xor data  '每一个数据与CRC寄存器进行异或
        For Flag = 0 To 7
          SaveHi = CRC16Hi
          SaveLo = CRC16Lo
          CRC16Hi = CRC16Hi \ 2            '高位右移一位
          CRC16Lo = CRC16Lo \ 2            '低位右移一位
          If ((SaveHi And &H1) = &H1) Then '如果高位字节最后一位为1
            CRC16Lo = CRC16Lo Or &H80      '则低位字节右移后前面补1
          End If                           '否则自动补0
          If ((SaveLo And &H1) = &H1) Then '如果LSB为1，则与多项式码进行异或
            CRC16Hi = CRC16Hi Xor CH
            CRC16Lo = CRC16Lo Xor CL
          End If
        Next Flag
     
 
    '  ReturnData(1) = CRC16Hi              'CRC高位
    '  ReturnData(0) = CRC16Lo              'CRC低位
 
      '低位在前，高位在后
  End Function
    
   'CRC高位字节值表
    Function GetCRCh(Ind As Byte) As Byte
      GetCRCh = Choose(Ind + 1, 0, 16, 32, 48, 64, 80, 96, 112, 129, 145, 161, 177, 193, 209, 225, 241, 18, 2, 50, 34, 82, 66, 114, 98, 147, 131, 179, 163, _
                    211, 195, 243, 227, 36, 52, 4, 20, 100, 116, 68, 84, 165, 181, 133, 149, 229, 245, 197, 213, 54, 38, 22, 6, 118, 102, 86, _
                    70, 183, 167, 151, 135, 247, 231, 215, 199, 72, 88, 104, 120, 8, 24, 40, 56, 201, 217, 233, 249, 137, 153, 169, 185, 90, _
                    74, 122, 106, 26, 10, 58, 42, 219, 203, 251, 235, 155, 139, 187, 171, 108, 124, 76, 92, 44, 60, 12, 28, 237, 253, 205, _
                    221, 173, 189, 141, 157, 126, 110, 94, 78, 62, 46, 30, 14, 255, 239, 223, 207, 191, 175, 159, 143, 145, 129, 177, 161, _
                    209, 193, 241, 225, 16, 0, 48, 32, 80, 64, 112, 96, 131, 147, 163, 179, 195, 211, 227, 243, 2, 18, 34, 50, 66, 82, 98, 114, _
                    181, 165, 149, 133, 245, 229, 213, 197, 52, 36, 20, 4, 116, 100, 84, 68, 167, 183, 135, 151, 231, 247, 199, 215, 38, 54, _
                    6, 22, 102, 118, 70, 86, 217, 201, 249, 233, 153, 137, 185, 169, 88, 72, 120, 104, 24, 8, 56, 40, 203, 219, 235, 251, 139, _
                    155, 171, 187, 74, 90, 106, 122, 10, 26, 42, 58, 253, 237, 221, 205, 189, 173, 157, 141, 124, 108, 92, 76, 60, 44, 28, 12, _
                    239, 255, 207, 223, 175, 191, 143, 159, 110, 126, 78, 94, 46, 62, 14, 30)

    End Function

    'CRC低位字节值表
    Function GetCRCl(Ind As Byte) As Byte
      GetCRCl = Choose(Ind + 1, 0, 33, 66, 99, 132, 165, 198, 231, 8, 41, 74, 107, 140, 173, 206, 239, 49, 16, 115, 82, 181, 148, 247, 214, 57, 24, 123, 90, _
            189, 156, 255, 222, 98, 67, 32, 1, 230, 199, 164, 133, 106, 75, 40, 9, 238, 207, 172, 141, 83, 114, 17, 48, 215, 246, 149, _
            180, 91, 122, 25, 56, 223, 254, 157, 188, 196, 229, 134, 167, 64, 97, 2, 35, 204, 237, 142, 175, 72, 105, 10, 43, 245, _
            212, 183, 150, 113, 80, 51, 18, 253, 220, 191, 158, 121, 88, 59, 26, 166, 135, 228, 197, 34, 3, 96, 65, 174, 143, 236, _
            205, 42, 11, 104, 73, 151, 182, 213, 244, 19, 50, 81, 112, 159, 190, 221, 252, 27, 58, 89, 120, 136, 169, 202, 235, 12, _
            45, 78, 111, 128, 161, 194, 227, 4, 37, 70, 103, 185, 152, 251, 218, 61, 28, 127, 94, 177, 144, 243, 210, 53, 20, 119, _
            86, 234, 203, 168, 137, 110, 79, 44, 13, 226, 195, 160, 129, 102, 71, 36, 5, 219, 250, 153, 184, 95, 126, 29, 60, 211, _
            242, 145, 176, 87, 118, 21, 52, 76, 109, 14, 47, 200, 233, 138, 171, 68, 101, 6, 39, 192, 225, 130, 163, 125, 92, 63, 30, _
            249, 216, 187, 154, 117, 84, 55, 22, 241, 208, 179, 146, 46, 15, 108, 77, 170, 139, 232, 201, 38, 7, 100, 69, 162, 131, _
            224, 193, 31, 62, 93, 124, 155, 186, 217, 248, 23, 54, 85, 116, 147, 178, 209, 240)

    End Function
Public Function CAL_CRC(data As String) As Long
 Dim lenss, I As Long
 Dim bbdata As Byte
 Dim sstrff As String
     sstrff = ""
     CRC16Lo = &HFF
     CRC16Hi = &HFF
     ReDim sentda(262) As Byte
     For I = 1 To 261
        sstrff = Mid(data, I * 2 - 1, 2)
        bbdata = Val("&H" & sstrff)
        CRC_oneByte (bbdata)
        sentda(I - 1) = bbdata
     Next I
      CAL_CRC = CRC16Hi
      CAL_CRC = CAL_CRC * 256 + CRC16Lo
End Function
 Function crc(daichuli() As Byte, l As Integer)
 Dim crch As Byte, crcl As Byte, da As Byte, I As Integer
 crch = 0
 I = 0
 crcl = 0
 While I <> l
    da = crch
    crch = crcl
    crcl = 0
    da = da Xor daichuli(I)
    crch = crch Xor GetCRCh(da)
    crcl = crcl Xor GetCRCl(da)
    I = I + 1
 Wend
       
     ReturnData(0) = crch              'CRC高位
     ReturnData(1) = crcl              'CRC低位
     ' crc = ReturnData
 'crc=0'
'while(len--!=0) {
'da=(uchar) (crc/256)' /* 以8 位二进制数的形式暂存CRC 的高8 位 */
'crc<<=8' /* 左移8 位，相当于CRC 的低8 位乘以28 */
'crc^=crc_ta[da^*ptr]' /* 高8 位和当前字节相加后再查表求CRC ，再加上以前的CRC */
'ptr++'
 End Function
 Public Function checkcrc(mm() As Byte, I As Integer)               'i为数组长度
 Dim dd(1) As Integer
 Call crc(mm, I - 2)
 If ReturnData(0) = mm(I - 2) And ReturnData(1) = mm(I - 1) Then
 checkcrc = 1
 Else
 checkcrc = 0
 End If
 End Function
 Public Function getready(ee() As Byte, I As Integer)               'i为数组长度
 Call crc(ee, I - 2)
 ee(I - 2) = ReturnData(0)
 ee(I - 1) = ReturnData(1)
 End Function


' 一个字节转 16进制字符串
Public Function byte_to_oneNum(temp As Long) As String
   Dim ssst As String
   temp = temp Mod 256
   ssst = byte_to_hex(temp)
   ssst = Mid(ssst, 1, 1)
   byte_to_oneNum = ssst
End Function
' 一个字节 或 数 转只有一位字符串
Public Function byte_to_hex(temp As Long) As String
    temp = temp Mod 256
  If temp < 16 Then
   byte_to_hex = "0" & Hex(temp)
  Else
  byte_to_hex = Hex(temp)
  End If
End Function
'两个字节转 16进制字符串
Public Function Int_to_hex(temp As Long) As String
    temp = temp Mod 65536
  If temp < 0 Then
     temp = -temp
    If temp < 16 Then
     Int_to_hex = "000" & Hex(temp)
    Else
      If temp < 16 * 16 Then
          Int_to_hex = "00" & Hex(temp)
      Else
            If temp < 16 * 16 * 16 Then
              Int_to_hex = "0" & Hex(temp)
             Else
               Int_to_hex = Hex(temp)
            End If
      End If
    End If
  Else

  If temp < 16 Then
   Int_to_hex = "000" & Hex(temp)
  Else
    If temp < 16 * 16 Then
        Int_to_hex = "00" & Hex(temp)
    Else
          If temp < 16 * 16 * 16 Then
            Int_to_hex = "0" & Hex(temp)
           Else
             Int_to_hex = Hex(temp)
          End If
    End If
  End If
  
  End If

End Function
'两个字节转 16进制字符串 小端模式
Public Function Int_to_Intel_hex(temp As Long) As String
    temp = temp Mod 65536
    If temp < 16 * 16 Then
        Int_to_Intel_hex = byte_to_hex(temp) & "00"
    Else
        Int_to_Intel_hex = byte_to_hex(temp Mod 256) & byte_to_hex(Fix(temp / 256))
    End If
 
End Function


