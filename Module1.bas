Attribute VB_Name = "mdlMain"
Global Color1 As RGBColor, Color2 As RGBColor
Global NewColor As RGBColor
Type RGBColor
    R As Integer
    G As Integer
    B As Integer
End Type


Public Function GetColorFromLong(LongColor As String) As RGBColor
    'great function (i did myself)
    'No "byte" stuff - just VB functions
    Dim H As String
    H = Hex(LongColor) ' convert the long color into hex color
    Do While Len(H) < 6
        H = 0 & H       'add leading zeros if needed
    Loop
    'extract each color and convert it to decimal color (normal format)
    GetColorFromLong.R = Val("&h" & Right(H, 2))
    GetColorFromLong.G = Val("&h" & Mid(H, 3, 2))
    GetColorFromLong.B = Val("&h" & Left(H, 2))
End Function


Public Function Blend(Color1 As RGBColor, Color2 As RGBColor, EnlargeR, EnlargeG, EnlargeB) As RGBColor
    Blend.R = (Color1.R * 2 + Color2.R * EnlargeR) / 3
    Blend.G = (Color1.G * 2 + Color2.G * EnlargeG) / 3
    Blend.B = (Color1.B * 2 + Color2.B * EnlargeB) / 3
End Function
