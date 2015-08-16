Attribute VB_Name = "FootprintDrawFunctions"
Option Explicit

' 列挙型
Public Enum TextTypeEnum
    TextTypeReference = 0
    TextTypeValue = 1
    TextTypeUser = 2
End Enum

Public Enum PadTypeEnum
    PadTypeSmd = 0
    PadTypeConnector = 1
    PadTypeThruHole = 2
    PadTypeNonThruHole = 3
End Enum

Public Enum PadShapeEnum
    PadShapeRect = 0
    PadShapeOval = 1
    PadShapeTrapezoid = 2
End Enum

' 直線の描画
Public Function DrawLine(x1 As Double, _
                         y1 As Double, _
                         x2 As Double, _
                         y2 As Double, _
                         layer As String, _
                         width As Double)
   DrawLine = Form("fp_line", _
                    Form("start", Dim2Str(x1), Dim2Str(y1)), _
                    Form("end", Dim2Str(x2), Dim2Str(y2)), _
                    Form("layer", layer), _
                    Form("width", Dim2Str(width)))
End Function

' 弧の描画
Public Function DrawArc(x1 As Double, _
                        y1 As Double, _
                        x2 As Double, _
                        y2 As Double, _
                        angle As Double, _
                        layer As String, _
                        width As Double)
   DrawArc = Form("fp_arc", _
                   Form("start", Dim2Str(x1), Dim2Str(y1)), _
                   Form("end", Dim2Str(x2), Dim2Str(y2)), _
                   Form("angle", Dim2Str(angle)), _
                   Form("layer", layer), _
                   Form("width", Dim2Str(width)))
End Function

' 円の描画
Public Function DrawCircle(x1 As Double, _
                           y1 As Double, _
                           x2 As Double, _
                           y2 As Double, _
                           layer As String, _
                           width As Double)
   DrawCircle = Form("fp_circle", _
                     Form("center", Dim2Str(x1), Dim2Str(y1)), _
                     Form("end", Dim2Str(x2), Dim2Str(y2)), _
                     Form("layer", layer), _
                     Form("width", Dim2Str(width)))
End Function

' テキストの描画
Public Function DrawText(textType As TextTypeEnum, _
                          str As String, _
                          x As Double, y As Double, angle As Double, _
                          layer As String, _
                          isHidden As Boolean, _
                          thickness As Double, _
                          w As Double, h As Double, _
                          isItalic As Boolean)
    Dim typeStr As String
    Select Case (textType)
        Case TextTypeEnum.TextTypeReference
            typeStr = "reference"
        Case TextTypeEnum.TextTypeValue
            typeStr = "value"
        Case TextTypeEnum.TextTypeUser
            typeStr = "user"
    End Select
    
    Dim escapedStr As String
    escapedStr = """" & EscapeString(str) & """"
            
    Dim hidden As String
    If isHidden Then hidden = "hide"
    
    Dim italic As String
    If isItalic Then italic = "italic"
                          
    DrawText = Form("fp_text", _
                    typeStr, _
                    escapedStr, _
                    Form("at", Dim2Str(x), Dim2Str(y), Dim2Str(angle)), _
                    Form("layer", layer), _
                    hidden, _
                    Form("effects", _
                        Form("font", _
                            Form("size", Dim2Str(w), Dim2Str(h)), _
                            Form("thickness", Dim2Str(thickness)), _
                            italic)))
End Function


' 文字列定数配列がないための代替案(範囲が限られているので)
Private Function padType(i As PadTypeEnum) As String
    Dim tmp As String
    Select Case i
    Case PadTypeSmd
        tmp = "smd"
    Case PadTypeConnector
        tmp = "connect"
    Case PadTypeThruHole
        tmp = "thru_hole"
    Case PadTypeNonThruHole
        tmp = "np_thru_hole"
    End Select
    padType = tmp
End Function

Private Function padShape(i As PadShapeEnum) As String
    Dim tmp As String
    Select Case i
    Case PadShapeRect
        tmp = "rect"
    Case PadShapeOval
        tmp = "oval"
    Case PadShapeTrapezoid
        tmp = "trapezoid"
    End Select
    padShape = tmp
End Function

' パッドの描画
Public Function DrawPad(padNum As Long, _
                    padType As PadTypeEnum, _
                    padShape As PadShapeEnum, _
                    x As Double, y As Double, _
                    w As Double, h As Double, _
                    w_short As Double, isHorizontalTrapezoid As Boolean, _
                    holeW As Double, holeH As Double, _
                    holeOffsetX As Double, holeOffsetY As Double, _
                    layers As String, _
                    die_length As Double) As String
    ' パッド番号(負は未定義の""にする)
    Dim padNumStr As String
    If padNum < 0 Then
        padNumStr = """"""""
    Else
        padNumStr = str(padNum)
    End If
    
    ' パッドの種類
    Dim padTypeStr As String
    Dim requireHole As Boolean
    Select Case padType
        Case PadTypeSmd
            padTypeStr = "smd"
            requireHole = False
        Case PadTypeConnector
            padTypeStr = "connect"
            requireHole = False
        Case PadTypeThruHole
            padTypeStr = "thru_hole"
            requireHole = True
        Case PadTypeNonThruHole
            padTypeStr = "np_thru_hole"
            requireHole = True
    End Select

    ' パッドの形状
    Dim padShapeStr As String
    Select Case padShape
    Case PadShapeRect
        padShapeStr = "rect"
    Case PadShapeOval
        padShapeStr = "oval"
    Case PadShapeTrapezoid
        padShapeStr = "trapezoid"
    End Select
    
    ' 台形の場合のみ必要な短辺長情報
    Dim rectDeltaStr As String
    If padShape = PadShapeTrapezoid Then
        If isHorizontalTrapezoid Then
            rectDeltaStr = Form("rect_delta", Dim2Str(w_short), Dim2Str(0))
        Else
            rectDeltaStr = Form("rect_delta", Dim2Str(0), Dim2Str(w_short))
        End If
    Else
        rectDeltaStr = ""
    End If
    
    ' オフセット
    Dim offsetStr As String
    offsetStr = Form("offset", Dim2Str(holeOffsetX), Dim2Str(holeOffsetY))

    ' 穴情報
    Dim drillStr As String
    If requireHole Then
        drillStr = Form("drill", "oval", Dim2Str(holeW), Dim2Str(holeH), offsetStr)
    Else
        drillStr = Form("drill", offsetStr)
    End If
    
    'ダイ-パッド間距離情報(負の値は不正なため情報を載せない。デフォルトは0)
    Dim dieLengthStr As String
    If die_length < 0 Then
        dieLengthStr = ""
    Else
        dieLengthStr = Form("die_length", Dim2Str(die_length))
    End If
    
    DrawPad = Form("pad", padNumStr, padTypeStr, padShapeStr, _
                   Form("at", Dim2Str(x), Dim2Str(y)), _
                   Form("size", Dim2Str(w), Dim2Str(h)), _
                   rectDeltaStr, _
                   drillStr, _
                   Form("layers", layers), _
                   dieLengthStr)
End Function

