VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "color2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public button As CommandButton
Public r_slider As Slider
Public g_slider As Slider
Public b_slider As Slider


Private colorCode As Long
Private enable As Boolean

Public Sub setEnable(IsEnable)
    enable = IsEnable
End Sub

Public Function getEnable() As Boolean
    getEnable = enable
    
End Function


Public Sub SetItem(button As CommandButton, r_slider As Slider _
                   , g_slider As Slider, b_slider As Slider, Item As String, xml As msxml, label As label, parentNode As String)
    Set Me.button = button
    Set Me.r_slider = r_slider
    Set Me.g_slider = g_slider
    Set Me.b_slider = b_slider
    
  

    colorCode = CLng(xml.getXml(Item)(0))

    If xml.getXml(parentNode)(0) = "True" Then
        label.ForeColor = 255
    End If
    
End Sub


   
Public Function getColorCode() As String
    getColorCode = colorCode
End Function
Public Function getRGBToColorCode(r As Integer, g As Integer, b As Integer) As String
    getRGBToColorCode = conversionToColorCode(r, g, b)
End Function


Private Function conversionToColorCode(r As Integer, g As Integer, b As Integer) As Long
    Dim rr As Long
    Dim gg As Long
    Dim bb As Long
    
    
    
    ' 赤の成分を計算
    bb = CLng(b) * 65536
    
    ' 緑の成分を計算
    gg = CLng(g) * 256
    
    ' 青の成分を計算
    conversionToColorCode = bb + gg + CLng(r)
End Function


Public Function GetRGB() As Variant
    GetRGB = conversionToRGB
    
End Function

Public Sub SetRGB(r As Integer, g As Integer, b As Integer)
    
    colorCode = conversionToColorCode(r, g, b)
End Sub

  
Public Sub SetSlider()
    Dim arr As Variant
        
    arr = conversionToRGB
    r_slider.value = arr(0)
    g_slider.value = arr(1)
    b_slider.value = arr(2)
    
    
End Sub
Private Function conversionToRGB() As Variant

    Dim arr As Variant
    
    Dim b As Long, g As Long, r As Long
    
    
    ' 赤の成分を計算
    b = colorCode \ 65536
    
    ' 緑の成分を計算
    g = (colorCode \ 256) Mod 256
    
    ' 青の成分を計算
    r = colorCode Mod 256
    
    arr = Array(r, g, b)
    
    conversionToRGB = arr
    
End Function


Public Sub changeColorXml(parentNode As String, xml As msxml)
    Call xml.writeXmlColor(parentNode, "colorCode", colorCode)
End Sub

Public Sub changeColorStatusXml(parentNode As String, xml As msxml)
    Call xml.writeXmlEnableColor(parentNode, "enable", enable)
End Sub


