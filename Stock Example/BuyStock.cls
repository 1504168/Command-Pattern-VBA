VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BuyStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("CommandPattern.Command.ConcreteImplementation")
'@PredeclaredId
'
'@Description : So this is the concrete implementation of the Command(Order) . It will execute the real command.
'                        Later We will Isolate  Real method (Buy or Sell) into an interface and Google, Apple will be the concrete
'                        implementation of that Interface.

'@Receiver

Option Explicit


Implements IOrder

Private Type TBuyStock
    GivenStock As Stock
End Type

Private This As TBuyStock

'@Override
Private Sub IOrder_Execute()
    Debug.Print "Bought "; "Stock Name : " & This.GivenStock.StockName & "    Quantity : " & This.GivenStock.Quantity
End Sub

Public Property Get GivenStock() As Stock
    Set GivenStock = This.GivenStock
End Property

Public Property Set GivenStock(ByVal RHS As Stock)
    Set This.GivenStock = RHS
End Property

Public Function Create(RequestForStock As Stock) As BuyStock
    Dim CurrentBuyStock As BuyStock
    Set CurrentBuyStock = New BuyStock
    With CurrentBuyStock
        Set .GivenStock = RequestForStock
    End With
    Set Create = CurrentBuyStock
End Function
