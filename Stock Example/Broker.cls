VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Broker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("CommandPattern.Invoker")
'@Description : So basically broker take request(Buy or Selling request) and then it execute that request based on the request type.
'                        Like if it is a request for Buying type(BuyStock) then it will execute the Buy method(sub) of Stock class.
'                        But if it is a request of Sell type (SellStock) then it will execute the Sell method(sub) of Stock class.
'                        Stock class is acting like an agency or company which can sell or buy the actual stock(Think Google as a company)

Option Explicit
Private Type TBroker
    Orders As Collection
End Type

Private This As TBroker

Public Property Get Orders() As Collection
    Set Orders = This.Orders
End Property

Public Property Set Orders(ByVal RHS As Collection)
    Set This.Orders = RHS
End Property

Public Sub TakeOrder(GivenOrder As IOrder)
    This.Orders.Add GivenOrder
End Sub

Public Sub PlaceOrders()
    
    Dim CurrentOrder As IOrder
    For Each CurrentOrder In This.Orders
        CurrentOrder.Execute
    Next CurrentOrder
    
End Sub

Private Sub Class_Initialize()
    Set This.Orders = New Collection
End Sub

Private Sub Class_Terminate()
    Set This.Orders = Nothing
End Sub

