Attribute VB_Name = "User"
'Four terms is always associated with this command pattern :

'1.Command(IOrder interface) >> It knows about the receiver and invoke execute method of that receiver

'2. Receiver(BuyStock and SellStock) >> Which execute the real business logic.

'3. Invoker(Broker) >> An invoker object knows how to execute a command, and optionally does bookkeeping
'                                   about the command execution. The invoker does not know anything about a concrete command,
'                                   it knows only about the command interface.

'4. Client(User Module) >> Invoker object, command objects and receiver objects are held by a client object, the client
'                                         decides which receiver objects it assigns to the command objects, and which commands it assigns
'                                          to the invoker.The client decides which commands to execute at which points. To execute a command,
'                                           it passes the command object to the invoker object.

'What about Stock then? >> (That is the Parametrized value for the Receiver).


'@Folder("CommandPattern.Command.InvokerUser")
'@It can be anything(Userform, Sheet interface or any other class)

Option Explicit


Public Sub Runner()
    
    Dim StockHandler As Broker
    Set StockHandler = New Broker
    
    StockHandler.TakeOrder CreateABuyRequest
    StockHandler.TakeOrder CreateASellRequest
    
    'Execute the taken order.
    StockHandler.PlaceOrders
    
End Sub

Public Function CreateABuyRequest() As BuyStock
    Dim FirstStock As Stock
    Set FirstStock = Stock.Create("Apple", 200)
    Dim BuyRequest As BuyStock
    Set BuyRequest = BuyStock.Create(FirstStock)
    Set CreateABuyRequest = BuyRequest
End Function


Public Function CreateASellRequest() As SellStock
    Dim FirstStock As Stock
    Set FirstStock = Stock.Create("Google", 100)
    Dim SellRequest As SellStock
    Set SellRequest = SellStock.Create(FirstStock)
    Set CreateASellRequest = SellRequest
End Function
