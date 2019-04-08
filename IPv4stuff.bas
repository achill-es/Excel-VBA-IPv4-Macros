Public Function IP2Int(ByVal IP2Int_s As String) As Double

    ' IPv4 addresses are strings that match the following EBNF spec:
    ' segment    : 0
    '            | [1-9][0-9]*
    '            ;
    ' ip-address : segment '.' segment '.' segment '.' segment
    '            ;

    ' this can also be specified as a regexp and processed via a finite automata:
    ' (0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)

    ' returns IP address as number or -1, if no match or illegal number

    IP2Int = -1

    Dim IP2Int_byte, IP2Int_ip As Double
    Dim IP2Int_i, IP2Int_dot, IP2Int_state As Byte
    Dim IP2Int_char As String

    IP2Int_byte = 0
    IP2Int_ip = 0
    IP2Int_dot = 0
    IP2Int_state = 0

    For IP2Int_i = 1 To Len(IP2Int_s)

        IP2Int_char = Mid(IP2Int_s, IP2Int_i, 1)

        Select Case IP2Int_char

        Case "0"
            Select Case IP2Int_state
            Case 0
                IP2Int_state = 1
            Case 1
                Exit Function
            'Case 2
                'IP2Int_state = 2
            End Select
            IP2Int_byte = IP2Int_byte * 10
            If IP2Int_byte > 255 Then Exit Function

        Case "1" To "9"
            If IP2Int_state = 1 Then
                Exit Function
            Else
                IP2Int_state = 2
            End If
            IP2Int_byte = IP2Int_byte * 10 + Val(IP2Int_char)
            If IP2Int_byte > 255 Then Exit Function

        Case "."
            If IP2Int_state = 0 Then
                Exit Function
            Else
                IP2Int_state = 0
            End If
            IP2Int_ip = IP2Int_ip * 256 + IP2Int_byte
            IP2Int_byte = 0
            IP2Int_dot = IP2Int_dot + 1

        Case Else
            Exit Function

        End Select

    Next IP2Int_i

    If (IP2Int_dot <> 3) Or (IP2Int_state = 0) Then Exit Function

    IP2Int = IP2Int_ip * 256 + IP2Int_byte

End Function

Public Function IsIPaddress(ByVal IsIPaddress_s As String) As Boolean

    Dim IsIPaddress_ip As Double

    IsIPaddress_ip = IP2Int(IsIPaddress_s)

    IsIPaddress = (0 <= IsIPaddress_ip) And (IsIPaddress_ip <= 4294967295#)

End Function

Public Function IsPublicIPaddress(ByVal IPaddress_s As String) As Boolean

    Dim IPaddress_ip As Double

    IPaddress_ip = IP2Int(IPaddress_s)
    IsPublicIPaddress = False

    'private IP addresses can't be public
    If IsRFC1918address(IPaddress_s) Then
        Exit Function

    '224.0.0.0/4 and higher
    ElseIf IPaddress_ip > 3758096383# Then
        Exit Function

    '0.0.0.0/8
    ElseIf IPaddress_ip < 16777216# Then
        Exit Function

    End If

    Select Case IPaddress_ip

    '127.0.0.0/8
    Case 2130706432# To 2147483647#
        Exit Function

    '169.254.0.0/16
    Case 2851995648# To 2852061183#
        Exit Function

    '192.0.2.0/24
    Case 3221225984# To 3221226239#
        Exit Function

    Case Else
        IsPublicIPaddress = True

    End Select

End Function

Public Function IsRFC1918address(ByVal IsRFC1918address_s As String) As Boolean

    Select Case IP2Int(IsRFC1918address_s)

    '10.0.0.0/8
    Case 167772160 To 184549375
        IsRFC1918address = True

    '172.16.0.0/12
    Case 2886729728# To 2887778303#
        IsRFC1918address = True

    '192.168.0.0/16
    Case 3232235520# To 3232301055#
        IsRFC1918address = True

    Case Else
        IsRFC1918address = False
    End Select

End Function

Public Function IsIPpool(ByVal IsIPpool_start As String, ByVal IsIPpool_finish As String) As Boolean

    Dim IsIPpool_start_ip, IsIPpool_finish_ip As Double

    IsIPpool_start_ip = IP2Int(IsIPpool_start)
    IsIPpool_finish_ip = IP2Int(IsIPpool_finish)

    IsIPpool = (0 < IsIPpool_start_ip) And (IsIPpool_start_ip <= IsIPpool_finish_ip)

End Function

Public Function PoolSize(ByVal Pool_min As String, ByVal Pool_max As String) As Double

    Dim IP_min, IP_max As Double

    If Pool_min = "" Then
        If Pool_max = "" Then
            PoolSize = 0
        Else
            PoolSize = -1
        End If
    Else
        If Pool_max = "" Then
            PoolSize = 1
        Else
            IP_max = IP2Int(Pool_max)
            IP_min = IP2Int(Pool_min)
            If IP_max >= IP_min Then
                PoolSize = IP_max - IP_min + 1
            Else
                PoolSize = -1
            End If
        End If
    End If

End Function

Public Function Pools2Route(ByVal Pools2Route_min As Double, ByVal Pools2Route_max As Double) As String

    Dim Pools2Route_network As Double
    Dim Pools2Route_netmask As Integer

    If (Pools2Route_min < 1) Then
        Pools2Route = "invalid Min"
        Exit Function
    End If
    If (Pools2Route_max < Pools2Route_min) Then
        Pools2Route = "invalid Max"
        Exit Function
    End If

    Pools2Route_netmask = 32

    Do While (Pools2Route_netmask > 1) And (Pools2Route_min <> Pools2Route_max)
        Pools2Route_min = Fix(Pools2Route_min / 2)
        Pools2Route_max = Fix(Pools2Route_max / 2)
        Pools2Route_netmask = Pools2Route_netmask - 1
    Loop

    Do While (Pools2Route_netmask > 30)
        Pools2Route_min = Fix(Pools2Route_min / 2)
        Pools2Route_netmask = Pools2Route_netmask - 1
    Loop

    Pools2Route_network = Pools2Route_min * (2 ^ (32 - Pools2Route_netmask))

    Pools2Route = Int2IP(Pools2Route_network) & " /" & Str(Pools2Route_netmask)

End Function

Public Function PoolStr2Route(ByVal PoolStr2Route_min As String, ByVal PoolStr2Route_max As String) As String

    Dim IP_min, IP_max As Double

    If PoolStr2Route_min = "" Then
        If PoolStr2Route_max = "" Then
            PoolStr2Route = ""
        Else
            PoolStr2Route = "missing Min"
        End If
    Else
        If PoolStr2Route_max = "" Or PoolStr2Route_max = PoolStr2Route_min Then
            PoolStr2Route = "255.255.255.255 / 32"
        Else
            IP_max = IP2Int(PoolStr2Route_max)
            IP_min = IP2Int(PoolStr2Route_min)
            If (IP_min > 0) And (IP_max > IP_min) Then
                PoolStr2Route = Pools2Route(IP_min, IP_max)
            Else
                PoolStr2Route = "invalid Pool"
            End If
        End If
    End If

End Function

Public Function IsSubnetMask(ByVal IPaddress As String, ByVal SubnetMask As String) As Boolean

    Dim IsSubnetMask_address, IsSubnetMask_mask As Double
    Dim IsSubnetMask_address_2, IsSubnetMask_mask_2 As Double
    Dim IsSubnetMask_addressLSB, IsSubnetMask_maskLSB As Double
    Dim IsSubnetMask_state, IsSubnetMask_32 As Byte

    IsSubnetMask = False

    IsSubnetMask_address = IP2Int(IPaddress)
    IsSubnetMask_mask = IP2Int(SubnetMask)

    If (IsSubnetMask_address = -1) Or (IsSubnetMask_mask = -1) Then
        Exit Function
    End If

    IsSubnetMask_state = 0

    For IsSubnetMask_32 = 1 To 32

        IsSubnetMask_address_2 = Fix(IsSubnetMask_address / 2)
        IsSubnetMask_addressLSB = IsSubnetMask_address - IsSubnetMask_address_2 * 2
        IsSubnetMask_address = IsSubnetMask_address_2

        IsSubnetMask_mask_2 = Fix(IsSubnetMask_mask / 2)
        IsSubnetMask_maskLSB = IsSubnetMask_mask - IsSubnetMask_mask_2 * 2
        IsSubnetMask_mask = IsSubnetMask_mask_2

        If IsSubnetMask_state = 0 Then
            If IsSubnetMask_maskLSB = 0 Then
                If IsSubnetMask_addressLSB = 1 Then Exit Function
            Else
                IsSubnetMask_state = 1
            End If
        Else
            If IsSubnetMask_maskLSB = 0 Then Exit Function
        End If

    Next IsSubnetMask_32

    IsSubnetMask = True

End Function

Public Function IsSubnetMaskCIDR(ByVal IPaddressCIDR As String) As Byte
    Dim IsSubnetMask_address, IsSubnetMask_exp As Double
    Dim IsSubnetMask_pos As Byte
    Dim IsSubnetMask_maskCIDR As Byte

    IsSubnetMaskCIDR = 0    'means error by default

    IsSubnetMask_pos = InStr(IPaddressCIDR, "/")
    If (IsSubnetMask_pos = 0) Then Exit Function      'this can be no CIDR notation

    IsSubnetMask_address = IP2Int(Left(IPaddressCIDR, IsSubnetMask_pos - 1))
    If (IsSubnetMask_address = -1) Then Exit Function

    IsSubnetMask_maskCIDR = Val(Right(IPaddressCIDR, Len(IPaddressCIDR) - IsSubnetMask_pos))
    If (IsSubnetMask_maskCIDR < 1) Or (IsSubnetMask_maskCIDR > 32) Then Exit Function

    IsSubnetMask_exp = 2 ^ (32 - IsSubnetMask_maskCIDR)

    If (IsSubnetMask_address = (Fix(IsSubnetMask_address / IsSubnetMask_exp) * IsSubnetMask_exp)) Then
        IsSubnetMaskCIDR = IsSubnetMask_maskCIDR
    End If

End Function

Public Function IsInsideIPpool(ByVal IsInsideIPpool_start As String, _
    ByVal IsInsideIPpool_finish As String, ByVal IsInsideIPpool_address As String) As Boolean

    Dim IsInsideIPpool_start_ip, IsInsideIPpool_finish_ip As Double
    Dim IsInsideIPpool_address_ip As Double

    IsInsideIPpool_start_ip = IP2Int(IsInsideIPpool_start)
    IsInsideIPpool_finish_ip = IP2Int(IsInsideIPpool_finish)
    IsInsideIPpool_address_ip = IP2Int(IsInsideIPpool_address)

    IsInsideIPpool = _
        IsIPpool(IsInsideIPpool_start_ip, IsInsideIPpool_finish_ip) And _
        (IsInsideIPpool_start_ip <= IsInsideIPpool_address_ip) And _
        (IsInsideIPpool_address_ip <= IsInsideIPpool_finish_ip)

End Function

Public Function Broadcast(ByVal IPaddress As String, ByVal SubnetMask As String) As String

    If IsSubnetMask(IPaddress, SubnetMask) Then
        Broadcast = Int2IP(IP2Int(IPaddress) + 4294967295# - IP2Int(SubnetMask))
    Else
        Broadcast = "invalid Mask"
    End If

End Function

Public Function XORed(ByVal SubnetMask As String) As String

    XORed = Int2IP(4294967295# - IP2Int(SubnetMask))

End Function

Public Function Int2IP(ByVal Int2IP_address As Double) As String

    Dim Int2IP_seg1 As Double
    Dim Int2IP_seg2, Int2IP_seg3, Int2IP_seg4 As Integer

    Int2IP_seg1 = Fix(Int2IP_address / 256)
    Int2IP_seg4 = Int2IP_address - Int2IP_seg1 * 256
    Int2IP_address = Int2IP_seg1

    Int2IP_seg1 = Fix(Int2IP_address / 256)
    Int2IP_seg3 = Int2IP_address - Int2IP_seg1 * 256
    Int2IP_address = Int2IP_seg1

    Int2IP_seg1 = Fix(Int2IP_address / 256)
    Int2IP_seg2 = Int2IP_address - Int2IP_seg1 * 256

    Int2IP = LTrim(Str(Int2IP_seg1)) & "." & LTrim(Str(Int2IP_seg2)) & "." _
            & LTrim(Str(Int2IP_seg3)) & "." & LTrim(Str(Int2IP_seg4))

End Function

Public Function IsSubnetHost(ByVal Network As String, ByVal SubnetMask As String, ByVal IPaddress As String) As Boolean

    Dim IsSubnetHost_net, IsSubnetHost_add As Double

    IsSubnetHost_net = IP2Int(Network)
    IsSubnetHost_add = IP2Int(IPaddress)

    IsSubnetHost = IsSubnetMask(Network, SubnetMask) And (IsSubnetHost_net < IsSubnetHost_add) And _
        (IsSubnetHost_add < IsSubnetHost_net + 4294967295# - IP2Int(SubnetMask))

End Function

Public Function NumberHosts(ByVal SubnetMask As String) As Double

    Dim NumberHosts_net, NumberHosts_net_2 As Double
    Dim NumberHosts_power As Byte

    NumberHosts_net = IP2Int(SubnetMask)
    NumberHosts_power = 0
    NumberHosts = 1

    Do
        NumberHosts_net_2 = Fix(NumberHosts_net / 2)

        If NumberHosts_net > NumberHosts_net_2 * 2 Then Exit Do

        NumberHosts_power = NumberHosts_power + 1
        NumberHosts_net = NumberHosts_net_2
        NumberHosts = NumberHosts * 2
    Loop

    If NumberHosts_power > 1 Then NumberHosts = NumberHosts - 2

End Function
