' Title: iLogic 規則是否可以引用部件中的元件並使用 Component.IsActive = True 或 False 控制其抑制狀態？- Can an iLogic rule reference a component in an assembly and control its suppression state using Component.IsActive = True or False?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-gui-ze-shi-fou-ke-yi-yin-yong-bu-jian-zhong-de-yuan-jian-bing-shi-yong-component-isactive-true-huo-false-kong-zhi-qi-yi-zhi-zhuang-tai/td-p/13800355
' Category: advanced
' Scraped: 2025-10-09T09:06:06.772814

Dim Group_DOOR_上 As Boolean = False
Dim Group_DOOR_下 As Boolean = False
Dim Group_DOOR_左 As Boolean = False
Dim Group_DOOR_右 As Boolean = False
Dim Group_SMIF_上 As Boolean = False
Dim Group_SMIF_下 As Boolean = False
Dim Group_SMIF_左 As Boolean = False
Dim Group_SMIF_右 As Boolean = False
Dim Group_LoadPort_上 As Boolean = False
Dim Group_LoadPort_下 As Boolean = False
Dim Group_LoadPort_左 As Boolean = False
Dim Group_LoadPort_右 As Boolean = False
Dim Group_Aligner_上 As Boolean = False
Dim Group_Aligner_下 As Boolean = False
Dim Group_Aligner_左 As Boolean = False
Dim Group_Aligner_右 As Boolean = False

' 根據您的選擇，獨立判斷並設定每個元件的狀態
' LoadPort
If 上 = "LoadPort" Then
    Group_LoadPort_上 = True
End If
If 下 = "LoadPort" Then
    Group_LoadPort_下 = True
End If
If 左 = "LoadPort" Then
    Group_LoadPort_左 = True
End If
If 右 = "LoadPort" Then
    Group_LoadPort_右 = True
End If
    
' DOOR
If 上 = "DOOR" Then
    Group_DOOR_上 = True    
End If
If 下 = "DOOR" Then	
    Group_DOOR_下 = True
End If
If 左 = "DOOR" Then
    Group_DOOR_左 = True
End If
If 右 = "DOOR" Then
    Group_DOOR_右 = True
End If
    
' SMIF
If 上 = "SMIF" Then
    Group_SMIF_上 = True
End If
If 下 = "SMIF" Then	
    Group_SMIF_下 = True
End If
If 左 = "SMIF" Then
    Group_SMIF_左 = True
End If
If 右 = "SMIF" Then
    Group_SMIF_右 = True
End If
    
' Aligner
If 上 = "Aligner" Then	
    Group_Aligner_上 = True
End If
If 下 = "Aligner" Then		
    Group_Aligner_下 = True
End If
If 左 = "Aligner" Then		
    Group_Aligner_左 = True
End If
If 右 = "Aligner" Then
    Group_Aligner_右 = True
End If

' 將變數值賦予元件，以啟用對應的設備

' LOAD PORT 元件啟用區塊
Component.IsActive("longup:1") = Group_LoadPort_上
Component.IsActive("shottup:1") =  Group_LoadPort_上
Component.IsActive("longdown:1") = Group_LoadPort_下
Component.IsActive("shotdown:1") =  Group_LoadPort_下
Component.IsActive("longleft:1") = Group_LoadPort_左
Component.IsActive("shotleft:1") =  Group_LoadPort_左
Component.IsActive("longright:1") = Group_LoadPort_右
Component.IsActive("shotright:1") =  Group_LoadPort_右

' DOOR
Component.IsActive("7DLR-0101-005:2") = Group_DOOR_上
Component.IsActive("7S022-0086-0101-038:2") = Group_DOOR_上
Component.IsActive("7S022-0086-0101-039:2") = Group_DOOR_上
Component.IsActive("AP-191-1-A:3") = Group_DOOR_上
Component.IsActive("AP-191-1-A:4") = Group_DOOR_上
Component.IsActive("7S022-0086-0101-057:2") = Group_DOOR_上
Component.IsActive("C-195-0H:3") = Group_DOOR_上
Component.IsActive("C-195-0H:4") = Group_DOOR_上
Component.IsActive("7DLR-0101-005:3") = Group_DOOR_下
Component.IsActive("7S022-0086-0101-038:3") = Group_DOOR_下
Component.IsActive("7S022-0086-0101-039:3") = Group_DOOR_下
Component.IsActive("AP-191-1-A:5") = Group_DOOR_下
Component.IsActive("AP-191-1-A:6") = Group_DOOR_下
Component.IsActive("7S022-0086-0101-057:3") = Group_DOOR_下
Component.IsActive("C-195-0H:5") = Group_DOOR_下
Component.IsActive("C-195-0H:6") = Group_DOOR_下
Component.IsActive("7DLR-0101-005:4") = Group_DOOR_左
Component.IsActive("7S022-0086-0101-038:4") = Group_DOOR_左
Component.IsActive("7S022-0086-0101-039:4") = Group_DOOR_左
Component.IsActive("AP-191-1-A:7") = Group_DOOR_左
Component.IsActive("AP-191-1-A:8") = Group_DOOR_左
Component.IsActive("7S022-0086-0101-057:4") = Group_DOOR_左
Component.IsActive("C-195-0H:7") = Group_DOOR_左
Component.IsActive("C-195-0H:8") = Group_DOOR_左
Component.IsActive("7DLR-0101-005:1") = Group_DOOR_右
Component.IsActive("7S022-0086-0101-038:1") = Group_DOOR_右
Component.IsActive("7S022-0086-0101-039:1") = Group_DOOR_右
Component.IsActive("AP-191-1-A:1") = Group_DOOR_右
Component.IsActive("AP-191-1-A:2") = Group_DOOR_右
Component.IsActive("7S022-0086-0101-057:1") = Group_DOOR_右
Component.IsActive("C-195-0H:1") = Group_DOOR_右
Component.IsActive("C-195-0H:2") = Group_DOOR_右
    
' SMIF
Component.IsActive("@3S025-9999-1601上:1") = Group_SMIF_上
Component.IsActive("@3S025-9999-1601下:1") = Group_SMIF_下
Component.IsActive("@3S025-9999-1601左:1") = Group_SMIF_左
Component.IsActive("@3S025-9999-1601右:1") = Group_SMIF_右

' Aligner
Component.IsActive("@LPT300-ST1 3D (RD-O3DR-24777)上:1") = Group_Aligner_上
Component.IsActive("@LPT300-ST1 3D (RD-O3DR-24777)下:1") = Group_Aligner_下
Component.IsActive("@LPT300-ST1 3D (RD-O3DR-24777)左:1") = Group_Aligner_左
Component.IsActive("@LPT300-ST1 3D (RD-O3DR-24777)右:1") = Group_Aligner_右