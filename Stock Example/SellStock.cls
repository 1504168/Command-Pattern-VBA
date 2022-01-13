VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SellStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("CommandPattern.Command.ConcreteImplementation")
'@PredeclaredId
'@Receiver

Option Explicit

Implements IOrder


Private Type TSellStock
    GivenStock As Stock
End Type

Private This As TSellStock

'@Override
Private Sub IOrder_Execute()
    Debug.Print "Sold "; "Stock Name : " & This.GivenStock.StockName & "    Quantity : " & This.GivenStock.Quantity
End Sub


Public Property Get GivenStock() As Stock
    Set GivenStock = This.GivenStock
End Property

Public Property Set GivenStock(ByVal RHS As Stock)
    Set This.GivenStock = RHS
End Property

Public Function Create(RequestForStock As Stock) As SellStock
    Dim CurrentSellStock As SellStock
    Set CurrentSellStock = New SellStock
    With CurrentSellStock
        Set .GivenStock = RequestForStock
    End With
    Set Create = CurrentSellStock
End Function

