input_mode (1)
ProgramName$ = "CI"
ProgramTitle$ = "Mobile POS"
VersionNumber$ = "1.2"

dbf% = 1

Comport% = 1
PrinterPort% = 1

Dim LangSet$(100)               		'Language Set Strings

Dim PrinterSet$(100) 	'Print out caption

Machine$ = GET_TARGET_MACHINE$
MaxLenBcode% = 16
MaxLenScode% = 16
MaxLenDesc% = 50
maxlencashinvoiceid% = 10
MaxLenQuantity% = 12
maxlenlocation% = 5
MaxLenPrice% = 10

gosub customize

'DBF 1
'Stockcode       16      Index
'Barcode          16      Index
'Description     50
'PriceA         10
'PriceB         10
'PriceC         10
'PriceD         10
'PriceE         10

'DBF2
'StockCode              16 Index
'LocationID             5 Index
'Quantity               10 Index

'DBF3
'CashInvoiceID  10      Index
'SubTotal               10
'Discount               10
'Total                  10
'Paid                   10
'Change         10
'Date Time              14

'DBF4
'CashInvoiceID  10
'ItemID           3
'Scode          16
'Desc                   50
'UnitPrice              10
'Discount               10
'Quantity               10
'Amount         10

if Machine$= "8300" then
	MaxLineChar% = 20
else
	Maxlinechar% = 16
End If

For Count = 1 To Maxlinechar%
        BlankLine$ = BlankLine$ + " "
Next

GoSub splashform


splashform:
                Cls
           locate 2, 1
           Print "  VIC COMPUTER  "
                locate 4, 1
                Print "   Mobile POS   "
                locate 5, 1
                Print "  Version " + VersionNumber$
                wait (300)
main:
	   GoSub SetLanguage
	  default_font%= 2
	   goSub SetPrinterLanguage
          	   GoSub mainmenu
                GoTo main

mainmenu:
	cursor(0)
                GoSub InitMenu
            Menu_str$ = LangSet$(75) + Chr$(13)
	Menu_str$ = Menu_str$ + LangSet$(1) + Chr$(13)
            Menu_str$ = Menu_str$ + LangSet$(2) + Chr$(13)
            Menu_str$ = Menu_str$ + LangSet$(3) + Chr$(13)
            Menu_str$ = Menu_str$ + LangSet$(4) + Chr$(13)
           Menu_str$ = Menu_str$ + LangSet$(5) + Chr$(13)
	Menu_str$ = Menu_str$ + "&7 Setup Mode" + chr(13)
	Menu_str$ = Menu_str$  + "&8 Customer" + chr(13)
           
MainMenuLoop:
             sel% = Menu(Menu_str$)
        If sel% = 1 Then
                GoSub InitInvoice
                GoSub ReadBcodeScreen
                locate 3, 1
                GoSub InputBcode
        else if sel% = 2 then
                GoSub CheckStockMenu
        else if sel% = 3 then
                GoSub TransferMenu
        else if sel% = 4 then
                GoSub statusmenu
        else if sel% = 5 then
                GoSub PrinterSetupMenu
        else if sel% = 6 then
                GoSub LanguageMenu
      else if sel% = 7 then
	   GoSub SetupMode
      else if sel% = 8 then
	   Gosub EnterCustomerCode
        Else
                GoTo MainMenuLoop
        End If
        GoTo MainMenuLoop

LanguageMenu:
        GoSub InitMenu
        Menu_str$ = LangSet$(6) + Chr$(13)
        Menu_str$ = Menu_str$ + LangSet$(7) + Chr$(13)
        Menu_str$ = Menu_str$ + LangSet$(8) + Chr$(13)

LanguageMenuLoop:
                sel% = Menu(Menu_str$)
        		On sel% GoSub SetEnglish, SetTradChin, SetSimpChin
                GoTo main
                return main

PrinterSetupMenu:
           GoSub InitMenu
            Menu_str$ = LangSet$(9) + Chr$(13)
        	Menu_str$ = Menu_str$ + LangSet$(10) + Chr$(13)
        	Menu_Str$ = Menu_Str$ + "&3 Language" + chr$(13)
	Menu_str$ = Menu_str$ + LangSet$(11) + Chr$(13)
        	Menu_str$ = Menu_str$ + LangSet$(12) + Chr$(13)
        	Menu_str$ = Menu_str$ + LangSet$(13) + Chr$(13)
	Menu_str$ = Menu_str$  + "&7 Drawer" + chr(13)

PrinterMenuLoop:
        sel% = Menu(Menu_str$)
        If sel% = 0 Then
                GoTo main
                return main
        End If
        On sel% GoSub EnablePrinterMenu, SelectPrinterMenu, PrinterLanguageMenu , PrintTest, SetHeader, SetFooter, EnableDrawerMenu
	GoTo PrinterSetupMenu

EnablePrinterMenu:
	GoSub InitMenu
	Menu_str$ = LangSet$(14) + Chr$(13)
	Menu_str$ = Menu_str$ + LangSet$(15) + Chr$(13)

EnablePrinterMenuLoop:
       sel% = Menu(Menu_str$)
        EnablePrinter% = sel%
        If EnablePrinter% = 1 Then
                locate 4, 1
                Print LangSet$(16)
                locate 6, 1
                Print LangSet$(17)
                wait (300)
        else if enableprinter% = 2 then
                locate 4, 1
                Print LangSet$(18)
                locate 6, 1
                Print LangSet$(19)
                wait (300)
                End If
        GoTo PrinterSetupMenu
                return PrinterSetupMenu

SelectPrinterMenu:
                GoSub InitMenu
	   Menu_str$ = "&1 TSP743" + Chr$(13)
                Menu_str$ = Menu_str$ +LangSet$(20) + Chr$(13)
        If Machine$ = "8060" Then
                Menu_str$ = Menu_str$ + LangSet$(21) + Chr$(13)
                Menu_str$ = Menu_str$ + LangSet$(22) + Chr$(13)
        End If

SelectPrinterMenuLoop:
	sel% = Menu(Menu_str$)
        	On sel% GoSub SetupStarSerial, SetupCitizenIR, SetupCitizenBT, SetupCustomBT
	Return

PrinterLanguageMenu:
                GoSub InitMenu
                Menu_str$ = LangSet$(6) + Chr$(13)
        	   Menu_str$ = Menu_str$ + LangSet$(7) + Chr$(13)
        	   Menu_str$ = Menu_str$ + LangSet$(8) + Chr$(13)

PrinterLanguageMenuLoop:
	  sel% = Menu(Menu_str$)
        	   On sel% GoSub SetEnglishPrinter, SetTradChinPrinter, SetSimpChinPrinter
                return

PrintTest:
        If Printer% < 1 Or Printer > 4 Then
                GoSub SelectPrinterMenu
        End If

        If Printer% = 1 Then
                
         	   GoSub ConnectSTARPrinter
	    Gosub TestSTARPrinter
	    Gosub DisconnectSTARPrinter

      else if printer% = 2 then

	    GoSub ConnectCitizenIRPrinter
                GoSub TestCitizenIRPrinter
           	     GoSub DisconnectCitizenIRPrinter

        else if printer% =3 then
                
                GoSub ConnectCitizenBTPrinter
                GoSub TestCitizenBTPrinter
                GoSub DisconnectCitizenBTPrinter

        else if printer% = 4 then
                
                GoSub ConnectCustomBTPrinter
                GoSub TestCustomBTPrinter
                GoSub DisconnectCustomBTPrinter
	     
      End If
        Return

SetupStarSerial:
      Printer% = 1
        EnablePrinter% = 1
        GoSub InitMenu
        locate 4, 1
        Print "Star serial"
        locate 6, 1
        Print "selected"
        wait (300)
        Return	

SetupCitizenIR:
        Printer% = 2
        EnablePrinter% = 1
        GoSub InitMenu
        locate 4, 1
        Print LangSet$(23)
        locate 6, 1
        Print LangSet$(24)
        wait (300)
        Return

SetupCitizenBT:
        GoSub BTPairing
        GoSub BTDeviceMenu
        GoSub InitMenu
        locate 4, 1
        If Len(DEVICE$) > 0 Then
                Printer% = 3
                EnablePrinter% = 1
                locate 4, 1
                Print LangSet$(25)
                locate 6, 1
                Print LangSet$(26)
        Else
                Print LangSet$(27)
        End If
        wait (300)
        Return

SetupCustomBT:
        GoSub BTPairing
        GoSub BTDeviceMenu
        GoSub InitMenu
        locate 4, 1
        If Len(DEVICE$) > 0 Then
                Printer% = 4
                EnablePrinter% = 1
                locate 4, 1
                Print LangSet$(28)
                locate 6, 1
                Print LangSet$(29)

        Else
                Print LangSet$(30)
        End If
        wait (300)
        Return


SetHeader:
       cursor(1)
        Select_font (1)
        locate 1, 1
        Print "Enter Header:"
        input header$
       cursor(0)
        return PrinterSetupMenu

SetFooter:
       cursor(1)
         Select_font (1)
        locate 1, 1
        Print "Enter Footer:"
        input footer$
       cursor(0) 
      return PrinterSetupMenu

EnableDrawerMenu:
        GoSub InitMenu
        Menu_str$ = "&1 Enable" + Chr$(13)
        Menu_str$ = Menu_str$ + "&2 Disable" + Chr$(13)

EnableDrawerMenuLoop:
    sel% = Menu(Menu_str$)
        EnableDrawer% = sel%
        If EnableDrawer% = 1 Then
                locate 4, 1
                Print "Cash drawer"
                locate 6, 1
                Print "enabled"
                wait (300)
        else if enabledrawer% = 2 then
                locate 4, 1
                Print "Cash drawer"
                locate 6, 1
                Print "disabled"
                wait (300)
        End If
        GoTo PrinterSetupMenu
        return PrinterSetupMenu

CheckStockMenu:
           GoSub InitMenu
           Menu_str$ = LangSet$(31) + Chr$(13)
           Menu_str$ = Menu_str$ + LangSet$(32) + Chr$(13)

CheckStockMenuLoop:
           sel% = Menu(Menu_str$)
            If sel% = 1 Then
                GoSub CheckPriceScreen
                GoSub CheckPrice
            else if sel% = 2 then
                GoSub CheckQuantityScreen
                GoSub CheckQuantity
           End If
        Return Main

CheckPriceScreen:
        GoSub InitMenu
        Select_font (2)
        locate 2, 1
        Print "Barcode:"
        Return

CheckPrice:
        On Reader(1) GoSub ReadPrice

        On Esc GoSub ExitReadPrice
        Enable Reader(1)
      cursor(1)
        input data$
	gosub customizeinput
      cursor(0)
      Off Esc
       Disable Reader(1)
       If Len(Data$) <= 0 Then
        GoSub ShowItemNotFound
        GoSub CheckPriceScreen
        GoTo CheckPrice
       Else
        GoSub FindPrice
        locate 8, 1
      End If
      Return Main

FindPrice:
       to_len% = MaxLenBcode%
        fillspacedata$ = Data$
        filldata$ = " "
        GoSub fill_space_500
        bcode$ = fillspacedata$

        If find_record(1, 2, bcode$) = 0 Then
           
        	GoSub ShowItemNotFound
	GoSub CheckPriceScreen

        Else
            
        	record$ = get_record(1, 2)
	gosub parserecord
        	Select_font (1)
        	locate 1, 1
        	Print scode$

	text$ = description$
	gosub gettextfont
	select_font(fonttype%)
	if fonttype% = 2 then
		locate 3,1
		print mid(text$,1,MaxChar% * 2)
	else
		locate 2,1
        		Print text$
	end if
	gosub clear4
	gosub clear5
	gosub clear6
	gosub clear7
	gosub clear8
	select_font(1)
	locate 4, 1
        	Print "A:" + Mid(record$, 83, 10)
        	locate 5, 1
        	Print "B:" + Mid(record$, 93, 10)
        	locate 6, 1
        	Print "C:" + Mid(record$, 103, 10)
        	locate 7, 1
        	Print "D:" + Mid(record$, 113, 10)
        	locate 8, 1
        	Print "E:" + Mid(record$, 123, 10)
        	locate 8, 1
      
        End If
        GoTo CheckPrice

ReadPrice:
	gosub readbarcode
        	Disable Reader(1)
	beep ( 3000,10 )
	return FindPrice
 
ExitReadPrice:
        Return Main


CheckQuantityScreen:
        GoSub InitMenu
        Select_font (2)
        locate 2, 1
        Print LangSet$(33)
        Return

CheckQuantity:
        On Reader(1) GoSub ReadQuantity
        On Esc GoSub ExitReadQuantity
       Enable Reader(1)
       cursor(1)
        input data$
	gosub customizeinput
       cursor(0) 
      Off Esc
       If Len(Data$) <= 0 Then
        GoSub ShowItemNotFound
        GoSub CheckQuantityScreen
        GoTo CheckQuantity
       Else
        GoSub FindQuantity
        locate 8, 1
      End If
      Return Main

FindQuantity:
       
        to_len% = MaxLenBcode%
        fillspacedata$ = Data$
        filldata$ = " "
        GoSub fill_space_500
        bcode$ = fillspacedata$
        If find_record(1, 2, bcode$) = 0 Then
                
        GoSub ShowItemNotFound
            GoSub CheckQuantityScreen

        Else
        record$ = get_record(1, 2)
        gosub parserecord

        Select_font (1)
        Cls
        locate 1, 1
        Print scode$

        text$ = description$
	gosub gettextfont
	select_font(fonttype%)
	if fonttype% = 2 then
		locate 3,1
		print mid(text$,1,MaxChar% * 2)
		gosub clear4
	else
		locate 2,1
        		Print text$
	end if
       select_font(1)
        Move_To(2,1)
      
        If find_record(2, 1, scode$) = 1 Then

                Menu_str$ = ""
                idx% = 0
	   locate 8,1
       	   print "Loading..."
                While (1 = 1)
                        record$ = get_record(2, 1)
                        If (Mid(record$, 1, 16) = scode$) Then
                                locate idx% + 4, 1
                                locationid$ = Mid(record$, 17, maxlenlocation%)
                                quantity$ = Mid(record$, 22, MaxLenQuantity%)
                                quantity$ = TRIM_RIGHT$(quantity$)
                                quantity$ = trim_left$(quantity$)
                                Print locationid$ + ": " + quantity$
                                idx% = idx% + 1
                        End If
                        If record_count(2) = get_record_number(2, 1) Then
                                exit
                        End If
                        Move_To_Next(2,1)
                Wend
	   gosub clear8
	   locate idx% + 4, 1
                                	   

        Else
                
                locate 5, 1
                Print "Empty"
        
        End If
             

      End If
        GoTo CheckQuantity

ReadQuantity:

       	gosub readbarcode
                beep ( 3000,10 )
                return FindQuantity

ExitReadQuantity:
        Disable Reader(1)
        Return Main

SelectPriceMenu:
           GoSub InitMenu
           Menu_str$ = LangSet$(34) + Chr$(13)
           Menu_str$ = Menu_str$ + LangSet$(35) + Chr$(13)
           Menu_str$ = Menu_str$ + LangSet$(36) + Chr$(13)
           Menu_str$ = Menu_str$ + LangSet$(37) + Chr$(13)
           Menu_str$ = Menu_str$ + LangSet$(38) + Chr$(13)
           Menu_str$ = Menu_str$ + LangSet$(39) + Chr$(13)
           Return

SelectPriceMenuLoop:
          sel% = Menu(Menu_str$)
          If sel% <> 0 And sel% < 6 Then
                PriceCur% = (sel% * MaxLenPrice%) + 73
                Price$ = Chr(64 + sel%)
                GoSub ParseRecord
                GoTo InputQuantity
         else if sel% = 6 then
                GoSub EnterUnitPrice
                GoTo InputQuantity
          Else
                GoTo main
          End If
             Return

EnterUnitPrice:
        to_len% = Maxlinechar% - 9
	fillspacedata$ = Str(unitprice!)
                filldata$ = ""
        leadingdata$ = " "
                GoSub fill_space_500
                PrintLine$ = fillspacedata$
        locate 2, 1
        Print LangSet$(40)
        locate 4, 1
        Print unitprice!
        locate 6, 1
        Print LangSet$(41)
        locate 8, 1
        Print "0"
        locate 8, 1
      cursor(1)
        input unitprice!
      cursor(0)
        Return

ExitReadBcode:
        Return Main

InitInvoice:
        quantity% = 0
        totalamount! = 0
        subtotal! = 0
        amount! = 0
        paidamount! = 0
        lastbcode$ = ""
        Price$ = "A"
        PriceCur% = 83
        EMPTY_TRANSACTION
        Return

Readbcode:

        gosub readbarcode
                Disable Reader(1)
                beep ( 3000,10 )
                return Findbcode

ReadBcodeScreen:
        GoSub InitMenu
        locate 2, 1
        Print LangSet$(42)
        locate 6, 1
        Print Str(subtotal!)
        locate 8, 1
        Print LangSet$(43) + Str(TRANSACTION_COUNT)
        Return

InputBcode:
        On Reader(1) GoSub Readbcode
        On Esc GoSub ExitReadBcode
        Enable Reader(1)
       locate 4, 1
      cursor(1)
        input data$
	gosub customizeinput
      cursor(0)
       Off Esc
       Disable Reader(1)
       If Len(Data$) <= 0 Then
        GoSub CheckOut
       Else
        GoSub FindBcode
      End If
      Return Main

datatobcode:
      to_len% = MaxLenBcode%
        fillspacedata$ = Data$
        filldata$ = " "
        GoSub fill_space_500
        bcode$ = fillspacedata$
       Return

FindBcode:
      GoSub datatobcode
      move_to(1,2,1)
        If find_record(1, 2, bcode$) = 0 Then
        locate 4, 1
            GoSub ShowItemNotFound
        GoSub ReadBcodeScreen
        return InputBcode
      Else
        record$ = get_record(1, 2)
                GoSub InitNewItem
        return inputquantity
     End If
      return  InputBcode

InitNewItem:
        
        GoSub ParseRecord
        quantity% = 1
        lastbcode$ = bcode$
        Return

ParseRecord:
        
        scode$ = Mid(record$, 1, MaxLenScode%)
        bcode$ = Mid(record$, 17, MaxLenBcode%)
      Description$ = Mid(record$, 33, MaxLenDesc%)
        unitprice! = VALR(Mid(record$, PriceCur%, MaxLenPrice%))
        Return

InputQuantity:
        Cls
        Select_font (1)
        locate 1, 1
        Print scode$
       text$ = description$
	gosub gettextfont
	select_font(fonttype%)
	if fonttype% = 2 then
		locate 3,1
		print mid(description$,1,MaxChar% * 2)
		gosub clear4
	else
		locate 2,1
        		Print mid(description$,1,MaxChar% * 3)

	end if
        Select_font (2)
        locate 6, 1
        Print Price$ + ": " + Str((unitprice! - unitdiscount!) * quantity%)
        locate 8, 1
        Print "Qty: " + Str(quantity%)
        Enable Reader(1)
        On Reader(1) GoSub ReadNextBcode
        On Esc GoSub ExitQuantity
        locate 8, 6
      cursor(1)
        input enterquantity%
      cursor(0)
        If enterquantity% <= 0 Then
                If quantity% <= 0 Then
                        quantity% = 1
                End If
        Else
                quantity% = enterquantity%
        End If
        
        GoSub SaveItem
        GoSub ReadBcodeScreen
        GoSub InputBcode

        Return

ExitQuantity:
        
        GoSub SelectPriceMenu
        Return SelectPriceMenuLoop

ReadNextBcode:
        
        gosub readbarcode
        Disable Reader(1)
        beep ( 3000,10 )
                GoSub datatobcode
        If lastbcode$ = bcode$ Then
                quantity% = quantity% + 1
                return InputQuantity
        End If
        
        If find_record(1, 2, bcode$) = 0 Then
                GoSub clear3
                GoSub clear4
                GoSub clear5
                GoSub clear6
                locate 4, 1
                GoSub ShowItemNotFound
                return InputQuantity
        Else
                bcode$ = lastbcode$
                GoSub SaveItem
                record$ = get_record(1, 2)
                        GoSub InitNewItem
                return InputQuantity
        End If
        return InputQuantity
                
SaveItem:
        to_len% = MaxLenQuantity%
                fillspacedata$ = Str(quantity%)
                filldata$ = " "
                GoSub fill_space_500
                quantity$ = fillspacedata$
        
        to_len% = MaxLenPrice%
                fillspacedata$ = Str(unitprice!)
                filldata$ = " "
                GoSub fill_space_500
                unitprice$ = fillspacedata$
        
        to_len% = MaxLenPrice%
                fillspacedata$ = Str(unitdiscount!)
                filldata$ = " "
                GoSub fill_space_500
                unitdiscount$ = fillspacedata$
        
        amount! = unitprice! * quantity%
        to_len% = MaxLenPrice%
                fillspacedata$ = Str(amount!)
                filldata$ = " "
                GoSub fill_space_500
                amount$ = fillspacedata$

        subtotal! = subtotal! + amount!
        transaction$ = scode$
        transaction$ = transaction$ + Description$
        transaction$ = transaction$ + unitprice$
        transaction$ = transaction$ + unitdiscount$
        transaction$ = transaction$ + quantity$
        transaction$ = transaction$ + amount$
        save_transaction (transaction$)
        Return

CheckOut:
        If TRANSACTION_COUNT <= 0 Then
                return inputbcode
        Else
                discount! = 0
                totalamount! = subtotal! - discount!
                GoSub EnterPaidAmount
        End If
        Return

ExitEnterPaidAmount:
        paidamount! = 0
        return EnterDiscountAmount

ExitEnterDiscountAmount:
        discount! = 0
        GoSub ReadBcodeScreen
        return inputBcode

checkoutscreen:

        to_len% = Maxlinechar% - 5
                fillspacedata$ = Str(subtotal!)
                filldata$ = ""
        leadingdata$ = " "
                GoSub fill_space_500
                PrintLine$ = fillspacedata$
        Subtotalline$ = "Sub: " + PrintLine$
        
        to_len% = Maxlinechar% - 6
                fillspacedata$ = Str(discount!)
                filldata$ = ""
        leadingdata$ = " "
                GoSub fill_space_500
                PrintLine$ = fillspacedata$
        Discountline$ = "Disc: " + PrintLine$
        
        to_len% = Maxlinechar% - 7
        totalamount! = subtotal! - discount!
        fillspacedata$ = Str(totalamount!)
                filldata$ = ""
        leadingdata$ = " "
                GoSub fill_space_500
                PrintLine$ = fillspacedata$
        Totalline$ = "Total: " + PrintLine$

        to_len% = Maxlinechar% - 6
                fillspacedata$ = Str(paidamount!)
                filldata$ = ""
        leadingdata$ = " "
                GoSub fill_space_500
                PrintLine$ = fillspacedata$
        PaidLine$ = "Paid: " + PrintLine$

        to_len% = Maxlinechar% - 8
                fillspacedata$ = Str(Change!)
                filldata$ = ""
        leadingdata$ = " "
                GoSub fill_space_500
                PrintLine$ = fillspacedata$
        changeline$ = "Change: " + PrintLine$
        Return

EnterDiscountAmount:

        totalamount! = subtotal!

        GoSub InitMenu
        GoSub checkoutscreen
        Select_font (1)
        locate 1, 1
        Print "   Check out    "

        locate 3, 1
        Print Subtotalline$
        
        locate 4, 1
        Print "Disc: 0"
        
        locate 5, 1
        Print Totalline$

        locate 6, 1
        Print PaidLine$

        locate 4, 7
        On Esc GoSub ExitEnterDiscountAmount
       cursor(1)
        input discount!
      cursor(0)
        
        If discount! > subtotal! Then
                locate 8, 1
                Print " Invalid Amount "
                wait (300)
                GoTo EnterDiscountAmount
        End If
        return EnterPaidAmount


EnterPaidAmount:
        GoSub InitMenu
        GoSub checkoutscreen
        Select_font (1)
        locate 1, 1
        Print "   Check out    "

        locate 3, 1
        Print Subtotalline$
        
        locate 4, 1
        Print Discountline$
        
        locate 5, 1
        Print Totalline$

        locate 6, 1
        Print "Paid: 0"

        locate 6, 7
        On Esc GoSub ExitEnterPaidAmount
       cursor(1)
        input paidamount!
        cursor(0)
      if paidamount! = 0 then
		paidamount! = totalamount!
      end if        


        GoSub checkoutscreen
                        
        locate 6, 1
        Print PaidLine$

        If paidamount! < totalamount! Then
                locate 8, 1
                Print " Insufficient "
                wait (300)
                GoTo EnterPaidAmount
        Else
                Change! = paidamount! - totalamount!
        End If
        
        GoSub SaveInvoice

        GoSub checkoutscreen

        locate 7, 1
        Print changeline$
        a$ = inkey$

      locate 8,1
     cursor(1)
      Print "ENTER to print"
    cursor(0)
        While (Asc(a$) <> 27) And (Asc(a$) <> 13)
                a$ = inkey$
        Wend
        printid$ = SerialID$
        If EnablePrinter% = 1 Then
	
 	  if Printer% = 1 then
		'STAR TSP700

		GoSub ConnectStarPrinter
		Gosub StarPrintInvoice
		Gosub OpenDrawer
		Gosub DisconnectStarPrinter


                else If Printer% = 2 Then
                        'Citizen IR Printer (CMP-10)
                
                        GoSub ConnectCitizenIRPrinter
                        GoSub CitizenIRPrintInvoice
                        GoSub DisconnectCitizenIRPrinter
                        
                else if printer%=3 then
                        
                        'Citizen IR Printer (CMP-10)

                        GoSub ConnectCitizenBTPrinter
                        GoSub CitizenIRPrintInvoice
                        GoSub DisconnectCitizenBTPrinter

                else if printer% =4 then
                        
                        'Custom BT Printer
                        GoSub ConnectCustomBTPrinter
                        GoSub CustomBTPrintInvoice
                        GoSub DisconnectCustomBTPrinter
                
                End If
        End If
        GoSub InitInvoice
        GoSub ReadBcodeScreen
        return inputbcode


OpenDrawer:
	If EnablePrinter% = 1 Then
		 if Printer% = 1 then
			write_com(printerport%,chr(7))
		 end if
	end if
	return

SaveInvoice:
        GoSub NextSerialID
        itemid% = 1

        to_len% = MaxLenPrice%
                fillspacedata$ = Str(subtotal!)
                filldata$ = " "
        leadingdata$ = ""
                GoSub fill_space_500
                subtotal$ = fillspacedata$

        to_len% = MaxLenPrice%
                fillspacedata$ = Str(discount!)
                filldata$ = " "
        leadingdata$ = ""
                GoSub fill_space_500
                discount$ = fillspacedata$
        
        to_len% = MaxLenPrice%
                fillspacedata$ = Str(totalamount!)
                filldata$ = " "
        leadingdata$ = ""
                GoSub fill_space_500
                totalamount$ = fillspacedata$

        to_len% = MaxLenPrice%
                fillspacedata$ = Str(paidamount!)
                filldata$ = " "
        leadingdata$ = ""
                GoSub fill_space_500
                paidamount$ = fillspacedata$
        
        to_len% = MaxLenPrice%
                fillspacedata$ = Str(Change!)
                filldata$ = " "
        leadingdata$ = ""
                GoSub fill_space_500
                Change$ = fillspacedata$

        record$ = SerialID$ + subtotal$ + discount$ + totalamount$ + paidamount$ + Change$ + Date$ + Time$
        
        add_record(3,record$)

        itemid% = 0
        While (transaction_count_ex(1) > 0)
                ItemData$ = get_transaction_data_ex(1, 1)
                itemid% = itemid% + 1
                to_len% = 3
                        fillspacedata$ = Str(itemid%)
                        filldata$ = " "
                leadingdata$ = ""
                        GoSub fill_space_500
                        itemid$ = fillspacedata$
                add_record(4,serialid$ + itemid$ + itemdata$)
                del_transaction_data_ex(1,1)
        Wend
        Return

NextSerialID:
        SerialID& = SerialID& + 1
        SerialID$ = Str(SerialID&)
        to_len% = maxlencashinvoiceid% - len(serialidprefix$)
        fillspacedata$ = SerialID$
        filldata$ = ""
        leadingdata$ = "0"
        GoSub fill_space_500
        SerialID$ = fillspacedata$
       SerialID$ =  serialidprefix$ + SerialID$
        Return

TestCitizenPrinter:
        Write_Com(PrinterPort%,chr(27) + chr(97)  + chr(1) + "VIC COMPUTER" + chr(10))
        Write_Com(PrinterPort%, chr(27) + chr(97)  + chr(1) + Ucase(ProgramTitle$) + chr(10))
        Write_Com(PrinterPort%, chr(27) + chr(97)  + chr(1) + "VER. " + VersionNumber$ + chr(10))
        Write_Com(PrinterPort%,chr(10))
        Write_Com(PrinterPort%, chr(27) + chr(97)  + chr(1) + "CITIZEN CMP-10" + chr(10))
        Write_Com(PrinterPort%, chr(27) + chr(97)  + chr(1) + "PRINT TEST" + chr(10))
        Write_Com(PrinterPort%,chr(10))
        Write_Com(PrinterPort%, chr(27) + chr(97)  + chr(1) + "HEADER" + chr(10))
        Write_Com(PrinterPort%,chr(27) + chr(33)  + chr(0) + "--------------------------------" + chr(10))
        If Len(Header$) > 0 Then
                Write_Com(PrinterPort%,chr(27) + chr(33)  + chr(0) + Header$ + chr(10))
        Else
                Write_Com(PrinterPort%,chr(27) + chr(33)  + chr(0) + "NULL" + chr(10))
        End If
        Write_Com(PrinterPort%,chr(27) + chr(33)  + chr(0) + "--------------------------------" + chr(10))
        Write_Com(PrinterPort%,chr(10))
        Write_Com(PrinterPort%, chr(27) + chr(97)  + chr(1) + "FOOTER" + chr(10))
        Write_Com(PrinterPort%,chr(27) + chr(33)  + chr(0) + "--------------------------------" + chr(10))
        If Len(footer$) > 0 Then
                Write_Com(PrinterPort%,chr(27) + chr(33)  + chr(0) + Footer$ + chr(10))
        Else
                Write_Com(PrinterPort%,chr(27) + chr(33)  + chr(0) + "NULL" + chr(10))
        End If
        Write_Com(PrinterPort%,chr(27) + chr(33)  + chr(0) + "--------------------------------" + chr(10))
        Write_Com(PrinterPort%,chr(10))
        Write_Com(PrinterPort%, chr(27) + chr(97)  + chr(1) + "DOWNLOADABLE GRAPHIC" + chr(10))
        Write_Com(PrinterPort%,chr(27) + chr(33)  + chr(0) + "--------------------------------" + chr(10))
        Write_Com(PrinterPort%,chr(29) + chr(47) + chr(0))
        Write_Com(PrinterPort%,chr(27) + chr(33)  + chr(0) + "--------------------------------" + chr(10))
        Write_Com(PrinterPort%,chr(10))
        Write_Com(PrinterPort%, chr(27) + chr(97)  + chr(1) + "END" + chr(10))
        Write_Com(PrinterPort%,chr(10) + chr(10) + chr(10)+ chr(10))
        Return

CitizenPrintInvoice:
        MaxLenPrintChar% = 32

        'Beep
        Write_Com(PrinterPort%, chr(27) + chr(30))
        
        Write_Com(PrinterPort%, chr(10) + chr(10))

        'Downloaded Graphic
        Write_Com(PrinterPort%,chr(29) + chr(47) + chr(0))

        If find_record(3, 1, printid$) = 1 Then

                record$ = get_record(3, 1)
                
                PrintLine$ = "   Sub total"
                PrintSubtotal$ = TRIM_RIGHT$(Mid(record$, 11, MaxLenPrice%))
                to_len% = MaxLenPrintChar% - Len(PrintSubtotal$)
                fillspacedata$ = PrintLine$
                filldata$ = " "
                GoSub fill_space_500
                PrintSubtotal$ = fillspacedata$ + PrintSubtotal$
                
                PrintLine$ = "   Discount"
                PrintDiscount$ = TRIM_RIGHT$(Mid(record$, 21, MaxLenPrice%))
                to_len% = MaxLenPrintChar% - Len(PrintDiscount$)
                fillspacedata$ = PrintLine$
                filldata$ = " "
                GoSub fill_space_500
                PrintDiscount$ = fillspacedata$ + PrintDiscount$
        
                PrintLine$ = "   Total"
                PrintTotalAmount$ = TRIM_RIGHT$(Mid(record$, 31, MaxLenPrice%))
                to_len% = MaxLenPrintChar% - Len(PrintTotalAmount$)
                fillspacedata$ = PrintLine$
                filldata$ = " "
                GoSub fill_space_500
                PrintTotalAmount$ = fillspacedata$ + PrintTotalAmount$

                PrintLine$ = "   Paid"
                PrintPaid$ = TRIM_RIGHT$(Mid(record$, 41, MaxLenPrice%))
                to_len% = MaxLenPrintChar% - Len(PrintPaid$)
                fillspacedata$ = PrintLine$
                filldata$ = " "
                GoSub fill_space_500
                PrintPaid$ = fillspacedata$ + PrintPaid$

                PrintLine$ = "   Change"
                PrintChange$ = TRIM_RIGHT$(Mid(record$, 51, MaxLenPrice%))
                to_len% = MaxLenPrintChar% - Len(PrintChange$)
                fillspacedata$ = PrintLine$
                filldata$ = " "
                GoSub fill_space_500
                PrintChange$ = fillspacedata$ + PrintChange$

                PrintCashInvoiceID$ = TRIM_RIGHT$(Mid(record$, 1, maxlencashinvoiceid%))
                
                PrintInvoiceDateTime$ = Mid(record$, 61, 4) + "."
                PrintInvoiceDateTime$ = PrintInvoiceDateTime$ + Mid(record$, 65, 2) + "."
                PrintInvoiceDateTime$ = PrintInvoiceDateTime$ + Mid(record$, 67, 2) + " "
                PrintInvoiceDateTime$ = PrintInvoiceDateTime$ + Mid(record$, 69, 2) + ":"
                PrintInvoiceDateTime$ = PrintInvoiceDateTime$ + Mid(record$, 71, 2)
        End If
        
        Write_Com(PrinterPort%, chr(10) + chr(10))
        
        If Len(Header$) > 0 Then
                Write_Com(PrinterPort%,chr(27) + chr(33)  + chr(0) + Header$ + chr(10))
        End If
        Write_Com(PrinterPort%, chr(27) + chr(97)  + chr(1)  + chr(27) + chr(33)  + chr(0) + "Cash Invoice" +  chr(10))
        Write_Com(PrinterPort%,chr(27) + chr(33)  + chr(0) + PrintCashInvoiceID$ + chr(10))
        Write_Com(PrinterPort%,chr(27) + chr(33)  + chr(0) +PrintInvoiceDateTime$ + chr(10))
        Write_Com(PrinterPort%,chr(27) + chr(33)  + chr(0) + "--------------------------------" + chr(10))

        BlankLine$ = ""

        Move_To(4,1)
        If find_record(4, 1, printid$) = 1 Then
                
                While (1 = 1)
                
                        ItemRecord$ = get_record(4, 1)

                        If (Mid(ItemRecord$, 1, maxlencashinvoiceid%) = printid$) Then
                                
                                itemcount% = itemcount% + 1
                                unitcount% = 0
                                PrintItemBcode$ = TRIM_RIGHT$(Mid(ItemRecord$, 14, MaxLenBcode%))
                                PrintItemDescription$ = " " + TRIM_RIGHT$(Mid(ItemRecord$, 30, 24))
                                PrintItemQuantity& = Val(Mid(ItemRecord$, 100, MaxLenQuantity%))
                                PrintItemAmount! = Val(Mid(ItemRecord$, 110, MaxLenQuantity%))
                                PrintItemUnitPrice! = VALR(Mid(ItemRecord$, 80, MaxLenPrice%))
                                Write_Com(PrinterPort%,Blankline$)
                                Write_Com(PrinterPort%,chr(27) + chr(33)  + chr(0) + PrintItemBcode$ + chr(10))
                                Write_Com(PrinterPort%,chr(27) + chr(33)  + chr(0) +PrintItemDescription$ + chr(10))

                                PrintLine$ = " " + Str(PrintItemQuantity&) + " x " + Str(PrintItemUnitPrice!)
                                PrintItemAmount! = Val(Mid(ItemRecord$, 110, MaxLenQuantity%))
                                to_len% = MaxLenPrintChar% - Len(Str(PrintItemAmount!))
                                fillspacedata$ = PrintLine$
                                filldata$ = " "
                                GoSub fill_space_500
                                PrintItemAmount$ = fillspacedata$ + Str(PrintItemAmount!)
                                
                                Write_Com(PrinterPort%,chr(27) + chr(33)  + chr(0) +PrintItemAmount$ + chr(10))
                                
                                BlankLine$ = Chr(10)

                        Else
        
                                exit
                        
                        End If

                        If record_count(4) = get_record_number(4, 1) Then
                                exit
                        End If

                        Move_To_Next(4,1)
        
                Wend

        End If

        Write_Com(PrinterPort%,chr(27) + chr(33)  + chr(0) + "--------------------------------" + chr(10))
        Write_Com(PrinterPort%,chr(27) + chr(33)  + chr(0) +PrintSubtotal$ + chr(10))
        Write_Com(PrinterPort%,chr(27) + chr(33)  + chr(0) + PrintDiscount$ + chr(10))
        Write_Com(PrinterPort%,chr(27) + chr(33)  + chr(0) +PrintTotalAmount$ + chr(10))
        Write_Com(PrinterPort%,chr(27) + chr(33)  + chr(0) +PrintPaid$ + chr(10))
        Write_Com(PrinterPort%,chr(27) + chr(33)  + chr(0) +PrintChange$ + chr(10))

        Write_Com(PrinterPort%, chr(10) + chr(10))
        
        Data$ = footer$
        Write_Com(PrinterPort%,chr(27) + chr(33)  + chr(0) + Data$ + chr(10))
        
        Write_Com(PrinterPort%, chr(10))
        Write_Com(PrinterPort%, chr(10))
        Write_Com(PrinterPort%, chr(10))
        Write_Com(PrinterPort%, chr(10))
                
        Return


ConnectStarPrinter:
      GoSub InitMenu
        locate 4, 1
        print "Connecting"
       locate 6,1 
      print "Star Printer"
      PrinterPort% = 1
      Set_Com_Type (PrinterPort%, 3)
      set_com (PrinterPort%, 5, 1, 2, 1 )                               ' com, 19200, n, 8, 1
        IRDA_TIMEOUT (1)
        OPEN_COM (PrinterPort%)
        MaxLenPrintChar% = 44
        While (IrDa_Status(0) <> 1)
                
        Wend


      return

TestStarPrinter:
        	Data$ =  "                VIC COMPUTER                " + chr(10)
       	gosub sendbuffer
        	Data$ = "                 " + Ucase(ProgramTitle$) + chr(10)
        	gosub sendbuffer
       	Data$ =   PrinterSet$(7)   + VersionNumber$ + chr(10)
         	gosub sendbuffer
	Data$ =chr(10)
 	gosub sendbuffer
        	Data$ = "                STAR TSP700                 " + chr(10)
         	gosub sendbuffer
	Data$ =   PrinterSet$(8)   + chr(10)
	gosub sendbuffer
        	Data$ =chr(10)
 	gosub sendbuffer
        	Data$ =   PrinterSet$(9)   + chr(10)
 	gosub sendbuffer
        	Data$ = "--------------------------------------------" + chr(10)
 	gosub sendbuffer
        	If Len(Header$) > 0 Then
                	Data$ = Header$ + chr(10)
 		gosub sendbuffer
        	Else
                	Data$ =   PrinterSet$(10)   + chr(10)
 		gosub sendbuffer
        	End If
        	Data$ = "--------------------------------------------" + chr(10)
 	gosub sendbuffer
        	Data$ =chr(10)
 	gosub sendbuffer
        	Data$ =   PrinterSet$(11)   + chr(10)
 	gosub sendbuffer
        	Data$ = "--------------------------------------------" + chr(10)
 	gosub sendbuffer
        	If Len(footer$) > 0 Then
                	Data$ =Footer$ + chr(10)
 		gosub sendbuffer
        	Else
                	Data$ =  PrinterSet$(12)   + chr(10)
 		gosub sendbuffer
        	End If
        	Data$ = "--------------------------------------------" + chr(10)
 	gosub sendbuffer
        	Data$ =chr(10)
 	gosub sendbuffer
        	Data$ =   PrinterSet$(13)   + chr(10)
 	gosub sendbuffer
        	Data$ = "--------------------------------------------" + chr(10)
 	gosub sendbuffer
        	'Data$ =chr(27) + chr(13) + chr(1)
 	'gosub sendbuffer
        	Data$ = "--------------------------------------------" + chr(10)
 	gosub sendbuffer
       	Data$ =chr(10)
 	gosub sendbuffer
        	Data$ =   PrinterSet$(14)   + chr(10)
 	gosub sendbuffer
        	Data$ =chr(10) + chr(10) + chr(10)+ chr(10) + chr(10) + chr(10) + chr(10) + chr(10)
	gosub sendbuffer
	Data$ = chr(27) + chr(100) + chr(1)
	gosub sendbuffer
	Data$ = chr(27) + chr(28) + "p" + chr(1) + chr(0)
      	gosub sendbuffer
	gosub SendBuffertoPrinter
        	Return

DisconnectStarPrinter:
      Set_Com_Type (PrinterPort%, 3)
        Close_Com (PrinterPort%)
      return

StarPrintInvoice:
	MaxLenPrintChar% = 44
        
        Data$ = chr(10) + chr(10)
		gosub sendbuffer
		
        If find_record(3, 1, printid$) = 1 Then

                record$ = get_record(3, 1)
                
                PrintLine$ =PrinterSet$(6)
                PrintSubtotal$ = TRIM_RIGHT$(Mid(record$, 11, MaxLenPrice%))
                to_len% = MaxLenPrintChar% - Len(PrintSubtotal$)
                fillspacedata$ = PrintLine$
                filldata$ = " "
                GoSub fill_space_500
                PrintSubtotal$ = fillspacedata$ + PrintSubtotal$
                
                PrintLine$ = PrinterSet$(1)
                PrintDiscount$ = TRIM_RIGHT$(Mid(record$, 21, MaxLenPrice%))
                to_len% = MaxLenPrintChar% - Len(PrintDiscount$)
                fillspacedata$ = PrintLine$
                filldata$ = " "
                GoSub fill_space_500
                PrintDiscount$ = fillspacedata$ + PrintDiscount$
        
                PrintLine$ =PrinterSet$(2)
                PrintTotalAmount$ = TRIM_RIGHT$(Mid(record$, 31, MaxLenPrice%))
                to_len% = MaxLenPrintChar% - Len(PrintTotalAmount$)
                fillspacedata$ = PrintLine$
                filldata$ = " "
                GoSub fill_space_500
                PrintTotalAmount$ = fillspacedata$ + PrintTotalAmount$

                PrintLine$ = PrinterSet$(3)
                PrintPaid$ = TRIM_RIGHT$(Mid(record$, 41, MaxLenPrice%))
                to_len% = MaxLenPrintChar% - Len(PrintPaid$)
                fillspacedata$ = PrintLine$
                filldata$ = " "
                GoSub fill_space_500
                PrintPaid$ = fillspacedata$ + PrintPaid$

                PrintLine$ = PrinterSet$(4)
                PrintChange$ = TRIM_RIGHT$(Mid(record$, 51, MaxLenPrice%))
                to_len% = MaxLenPrintChar% - Len(PrintChange$)
                fillspacedata$ = PrintLine$
                filldata$ = " "
                GoSub fill_space_500
                PrintChange$ = fillspacedata$ + PrintChange$

                PrintCashInvoiceID$ = TRIM_RIGHT$(Mid(record$, 1, maxlencashinvoiceid%))
                
                PrintInvoiceDateTime$ = Mid(record$, 61, 4) + "."
                PrintInvoiceDateTime$ = PrintInvoiceDateTime$ + Mid(record$, 65, 2) + "."
                PrintInvoiceDateTime$ = PrintInvoiceDateTime$ + Mid(record$, 67, 2) + " "
                PrintInvoiceDateTime$ = PrintInvoiceDateTime$ + Mid(record$, 69, 2) + ":"
                PrintInvoiceDateTime$ = PrintInvoiceDateTime$ + Mid(record$, 71, 2)

        End If
        
        Data$ = chr(10) + chr(10)
		gosub sendbuffer
        
        If Len(Header$) > 0 Then
                Data$ =  Header$ + chr(10)
				gosub sendbuffer
        End If
        Data$ = chr(27) + chr(97)  + chr(1)  +  PrinterSet$(5) +  chr(10)
		gosub sendbuffer
        Data$ =  PrintCashInvoiceID$ + chr(10)
		gosub sendbuffer
        Data$ = PrintInvoiceDateTime$ + chr(10)
		gosub sendbuffer
        Data$ = "--------------------------------------------" + chr(10)
		gosub sendbuffer

        BlankLine$ = ""

        Move_To(4,1)
        If find_record(4, 1, printid$) = 1 Then
                
                While (1 = 1)
                
                        ItemRecord$ = get_record(4, 1)

                        If (Mid(ItemRecord$, 1, maxlencashinvoiceid%) = printid$) Then
                                
                                itemcount% = itemcount% + 1
                                unitcount% = 0
                                PrintItemBcode$ = TRIM_RIGHT$(Mid(ItemRecord$, 14, MaxLenBcode%))
                                PrintItemDescription$ = " " + TRIM_RIGHT$(Mid(ItemRecord$, 30, 30))
                                PrintItemQuantity& = Val(Mid(ItemRecord$, 100, MaxLenQuantity%))
                                PrintItemAmount! = Val(Mid(ItemRecord$, 110, MaxLenQuantity%))
                                PrintItemUnitPrice! = VALR(Mid(ItemRecord$, 80, MaxLenPrice%))
                                Data$ = Blankline$
			gosub sendbuffer
                                Data$ =  PrintItemBcode$ + chr(10)
			gosub sendbuffer
                                Data$ = PrintItemDescription$ + chr(10)
			gosub sendbuffer
		      PrintLine$ = " " + Str(PrintItemQuantity&) + " x " + Str(PrintItemUnitPrice!)
                                PrintItemAmount! = Val(Mid(ItemRecord$, 110, MaxLenQuantity%))
                                to_len% = MaxLenPrintChar% - Len(Str(PrintItemAmount!))
                                fillspacedata$ = PrintLine$
                                filldata$ = " "
                                GoSub fill_space_500
                                PrintItemAmount$ = fillspacedata$ + Str(PrintItemAmount!)
                                
                                Data$ = PrintItemAmount$ + chr(10)
                                gosub sendbuffer
                                
                                BlankLine$ = Chr(10)

                        Else
        
                                exit
                        
                        End If

                        If record_count(4) = get_record_number(4, 1) Then
                                exit
                        End If

                        Move_To_Next(4,1)
        
                Wend

        End If

        Data$ = "--------------------------------------------" + chr(10)
		gosub sendbuffer
        Data$ =PrintSubtotal$ + chr(10)
        gosub sendbuffer
		Data$ =  PrintDiscount$ + chr(10)
        gosub sendbuffer
        Data$ = PrintTotalAmount$ + chr(10)
        gosub sendbuffer
        Data$ = PrintPaid$ + chr(10)
        gosub sendbuffer
        
	Data$ = PrintChange$ + chr(10)
	gosub sendbuffer
        	
 		
        Data$ = chr(10) + chr(10)
        gosub sendbuffer
        
        Data$ = footer$
        gosub sendbuffer
        
        Data$ =  chr(10)
        gosub sendbuffer
        
        Data$ = chr(10)
        gosub sendbuffer
        
        Data$ = chr(10)
        gosub sendbuffer
        
        Data$ = chr(10)
        gosub sendbuffer
        
        Data$ = chr(10)
        gosub sendbuffer

      Data$ = chr(27) + chr(100) + chr(1)
      gosub sendbuffer
	
     Data$ = chr(27) + chr(28) + "p" + chr(1) + chr(0)
      gosub sendbuffer
	
        gosub sendbuffertoprinter
                
        Return

SendBuffer:

	'To minimize the use of write_com
	if len(data$) + len(buffer$) > 255 then
		gosub sendbuffertoprinter
		buffer$ = data$
	else
		buffer$ = buffer$ + data$
	end if
	return
	

SendBufferToPrinter:
	senddata$ = buffer$
	buffer$ = ""
	Write_Com(PrinterPort%, senddata$)
	return


ConnectCitizenIRPrinter:
        GoSub InitMenu
        locate 4, 1
        Print LangSet$(44)
        locate 6, 1
        Print LangSet$(45)
        PrinterPort% = 1
        Set_Com_Type (PrinterPort%, 4)
        IRDA_TIMEOUT (1)
        OPEN_COM (PrinterPort%)
        MaxLenPrintChar% = 32
        While (IrDa_Status(0) <> 1)
                
        Wend

        Return

CitizenIRPrintInvoice:
        GoSub CitizenPrintInvoice
        Return

TestCitizenIRPrinter:
        GoSub TestCitizenPrinter
        Return

DisconnectCitizenIRPrinter:

        Set_Com_Type (PrinterPort%, 3)
        Close_Com (PrinterPort%)
        Return

ConnectCitizenBTPrinter:
        GoSub InitMenu
        locate 4, 1
        Print LangSet$(46)
        locate 6, 1
        Print LangSet$(47)
        PrinterPort% = 2

        'Set com 2 as Bluetooth com port (SPP)
        SET_COM_TYPE(PrinterPort%,5)

        'Set_COM(N%,Baudrate%, Parity%,Data%,Handshake%)

        'N% - 2 is for SPP
        'Baudrate% - 4 is Master for SPP
        'Parity% - Not applied in Bluetooth, Set to 1
        'Data% - Not applied in Bluetooth, Set to 1
        'Handshake% - Set 4 for Wedge Emulator
        SET_COM(2,4,1,2,1)
        OPEN_COM (PrinterPort%)
        
        While (GET_WLAN_STATUS(1) = 0)

        Wend
        BEEP(4400,4)
        

        Return

TestCitizenBTPrinter:
        GoSub TestCitizenPrinter
        Return
        
CitizenBTPrintInvoice:
        
        GoSub CitizenPrintInvoice
        Return

DisconnectCitizenBTPrinter:
        Close_Com (PrinterPort%)
        Return

ConnectCustomBTPrinter:
        GoSub InitMenu
        locate 4, 1
        Print LangSet$(48)
        locate 6, 1
        Print LangSet$(49)

        'Set com 2 as Bluetooth com port (SPP)
        SET_COM_TYPE(2,5)

        'Set_COM(N%,Baudrate%, Parity%,Data%,Handshake%)

        'N% - 2 is for SPP
        'Baudrate% - 4 is Master for SPP
        'Parity% - Not applied in Bluetooth, Set to 1
        'Data% - Not applied in Bluetooth, Set to 1
        'Handshake% - Set 4 for Wedge Emulator
        SET_COM(2,4,1,2,4)
        OPEN_COM (2)
        MaxLenPrintChar% = 24
        

WedgeLoop:
        If WEDGE_READY = 0 Then GoTo WedgeLoop
        BEEP(4400,4)
        Return

DisconnectCustomBTPrinter:
        Close_Com (PrinterPort%)
        Return

TestCustomBTPrinter:

        send_wedge (Chr(27) + Chr(78))
        send_wedge ("      VIC COMPUTER      ")
        send_wedge ("       MOBILE POS       ")
        send_wedge ("        VER. " + VersionNumber$ + Chr(10))
        send_wedge (Chr(10))
        send_wedge ("   CUSTOM S'PRINT-BT    ")
        send_wedge ("       PRINT TEST       ")
        send_wedge (Chr(10))
        send_wedge ("         HEADER         ")
        send_wedge ("------------------------")
        If Len(Header$) > 0 Then
                send_wedge (Header$)
        Else
                send_wedge ("NULL" + chr(10))
        End If
        send_wedge ("------------------------")
        send_wedge (Chr(10))
        send_wedge ("         FOOTER         ")
        send_wedge ("------------------------")
        If Len(footer$) > 0 Then
                send_wedge (footer$)
        Else
                send_wedge ("NULL" + chr(10))
        End If
        send_wedge ("------------------------")
        send_wedge (Chr(10))
        send_wedge ("  DOWNLOADABLE GRAPHIC  ")
        send_wedge ("------------------------")
        send_wedge (Chr(27) + Chr(250) + Chr(1) + Chr(255) + Chr(10))
        send_wedge ("------------------------")
        send_wedge (Chr(10))
        send_wedge ("          END           ")
        send_wedge (Chr(10) + Chr(10) + Chr(10) + Chr(10))
        Return

CustomBTPrintInvoice:
                
        MaxLenPrintChar% = 24

        send_wedge (Chr(27) + Chr(78))
         'Clear Printer buffer
                
        send_wedge (Chr(7))
                'Downloadable Graphic
                send_wedge (Chr(27) + Chr(250) + Chr(1) + Chr(255))
        
        send_wedge (Chr(10) + Chr(10))

        If find_record(3, 1, printid$) = 1 Then

                record$ = get_record(3, 1)
                
                PrintLine$ = "   Sub total"
                PrintSubtotal$ = TRIM_RIGHT$(Mid(record$, 11, MaxLenPrice%))
                to_len% = MaxLenPrintChar% - Len(PrintSubtotal$)
                fillspacedata$ = PrintLine$
                filldata$ = " "
                GoSub fill_space_500
                PrintSubtotal$ = fillspacedata$ + PrintSubtotal$
                
                PrintLine$ = "   Discount"
                PrintDiscount$ = TRIM_RIGHT$(Mid(record$, 21, MaxLenPrice%))
                to_len% = MaxLenPrintChar% - Len(PrintDiscount$)
                fillspacedata$ = PrintLine$
                filldata$ = " "
                GoSub fill_space_500
                PrintDiscount$ = fillspacedata$ + PrintDiscount$
        
                PrintLine$ = "   Total"
                PrintTotalAmount$ = TRIM_RIGHT$(Mid(record$, 31, MaxLenPrice%))
                to_len% = MaxLenPrintChar% - Len(PrintTotalAmount$)
                fillspacedata$ = PrintLine$
                filldata$ = " "
                GoSub fill_space_500
                PrintTotalAmount$ = fillspacedata$ + PrintTotalAmount$

                PrintLine$ = "   Paid"
                PrintPaid$ = TRIM_RIGHT$(Mid(record$, 41, MaxLenPrice%))
                to_len% = MaxLenPrintChar% - Len(PrintPaid$)
                fillspacedata$ = PrintLine$
                filldata$ = " "
                GoSub fill_space_500
                PrintPaid$ = fillspacedata$ + PrintPaid$

                PrintLine$ = "   Change"
                PrintChange$ = TRIM_RIGHT$(Mid(record$, 51, MaxLenPrice%))
                to_len% = MaxLenPrintChar% - Len(PrintChange$)
                fillspacedata$ = PrintLine$
                filldata$ = " "
                GoSub fill_space_500
                PrintChange$ = fillspacedata$ + PrintChange$

                PrintCashInvoiceID$ = TRIM_RIGHT$(Mid(record$, 1, maxlencashinvoiceid%))
                
                PrintInvoiceDateTime$ = Mid(record$, 61, 4) + "."
                PrintInvoiceDateTime$ = PrintInvoiceDateTime$ + Mid(record$, 65, 2) + "."
                PrintInvoiceDateTime$ = PrintInvoiceDateTime$ + Mid(record$, 67, 2) + " "
                PrintInvoiceDateTime$ = PrintInvoiceDateTime$ + Mid(record$, 69, 2) + ":"
                PrintInvoiceDateTime$ = PrintInvoiceDateTime$ + Mid(record$, 71, 2)
        End If
        
        send_wedge (Header$)
        
        send_wedge (Chr(10))

        send_wedge ("      Cash Invoice      ")
        send_wedge (PrintCashInvoiceID$ + Chr(10))
        send_wedge (PrintInvoiceDateTime$ + Chr(10))
        
        send_wedge ("------------------------")

        
        BlankLine$ = ""

        Move_To(4,1)
        If find_record(4, 1, printid$) = 1 Then
        

                While (1 = 1)
                
                        ItemRecord$ = get_record(4, 1)

                        If (Mid(ItemRecord$, 1, maxlencashinvoiceid%) = printid$) Then
                                
                                itemcount% = itemcount% + 1
                                unitcount% = 0
                                PrintItemBcode$ = TRIM_RIGHT$(Mid(ItemRecord$, 14, MaxLenBcode%)) + Chr(10)
                                PrintItemDescription$ = " " + TRIM_RIGHT$(Mid(ItemRecord$, 30, 16)) + Chr(10)
                                PrintItemQuantity& = Val(Mid(ItemRecord$, 100, MaxLenQuantity%))
                                PrintItemAmount! = Val(Mid(ItemRecord$, 110, MaxLenQuantity%))
                                PrintItemUnitPrice! = VALR(Mid(ItemRecord$, 80, MaxLenPrice%))
                                send_wedge (BlankLine$)
                                send_wedge (PrintItemBcode$)
                                send_wedge (PrintItemDescription$)

                                PrintLine$ = " " + Str(PrintItemQuantity&) + " x " + Str(PrintItemUnitPrice!)
                                PrintItemAmount! = Val(Mid(ItemRecord$, 110, MaxLenQuantity%))
                                to_len% = MaxLenPrintChar% - Len(Str(PrintItemAmount!))
                                fillspacedata$ = PrintLine$
                                filldata$ = " "
                                GoSub fill_space_500
                                PrintItemAmount$ = fillspacedata$ + Str(PrintItemAmount!)
                                
                                send_wedge (PrintItemAmount$)
                                
                                BlankLine$ = Chr(10)

                        Else
        
                                exit
                        
                        End If

                        If record_count(4) = get_record_number(4, 1) Then
                                exit
                        End If

                        Move_To_Next(4,1)
        
                Wend

        End If

        send_wedge ("------------------------")
        send_wedge (PrintSubtotal$)
        send_wedge (PrintDiscount$)
        send_wedge (PrintTotalAmount$)
        send_wedge (PrintPaid$)
        send_wedge (PrintChange$)

        send_wedge (Chr(10) + Chr(10))
        
        Data$ = footer$
        GoSub sendslice

        send_wedge (Chr(10) + Chr(10) + Chr(10) + Chr(10))

        Return


sendslice:
        For idx% = 0 To Len(Data$) Step 128
                senddata$ = Mid(Data$, idx% + 1, 128)
                send_wedge (senddata$)
        Next
        Return

TransferMenu:

        If Machine$ = "8000" Then
                GoSub transferdata
    

        else if Machine$="8060" then

                Menu_str$ = LangSet$(50) + Chr$(13)
                Menu_str$ = Menu_str$ + LangSet$(51) + Chr$(13)
                sel% = Menu(Menu_str$)
                If sel% = 1 Then
                        GoSub transferdata
                else if sel% = 2 then
                        GoSub BTTransfer
                Else
                        return main
                End If

      else if Machine$ = "8300" then
	   
	   Menu_str$ = "1. IR Cradle" + Chr$(13)
                Menu_str$ = Menu_Str$ + "2. RS232" + Chr$(13)
                sel% = Menu(Menu_str$)
                If sel% = 1 Then
		ComPort% = 1
                        GoSub transferdata
                else if sel% = 2 then
		ComPort% = 1
                        GoSub Enable_IRCradle
                Else
                        return main
                End If

      else

		gosub transferdata

        End If
        return main

BTPairing:

        Cls
        locate 3, 1
        Print LangSet$(52)
        locate 5, 1
        Print LangSet$(53)
        BTDeviceStr$ = BT_INQUIRY$
        GoSub ParseBTDeviceStr
        Return

BTDeviceMenu:
        i% = 0
        i% = Menu(BTDeviceNameStr$)
        DEVICE$ = ""
        If i% > 0 Then
                DEVICE$ = Mid$(BTDeviceStr$, 1 + 32 * (i% - 1), 32)
                Cls
                locate 4, 1
                Print LangSet$(54)
                locate 6, 1
                Print Mid$(DEVICE$, 13, 20)
                i% = BT_PAIRING(DEVICE$, 3)
                If i% <> 1 Then
                        BEEP(3300,4)
                        Print LangSet$(55)
                        wait (400)
                        GoTo BTDeviceMenu
                End If
        End If
        Return

ParseBTDeviceStr:
        BTDeviceNameStr$ = ""
        i% = 1
        For idx% = 1 To Len(BTDeviceStr$) Step 32
                DeviceName$ = Mid(BTDeviceStr$, (idx% - 1) + 13, 20)
                t$ = Chr$(13)
                l% = InStr(DeviceName$, t$)
                If i% > 0 Then
                        DeviceName$ = Mid(DeviceName$, 1, l% - 1)
                End If
                BTDeviceNameStr$ = BTDeviceNameStr$ + Mid("&" + Str(i%) + " " + DeviceName$, 1, 11) + Chr(13)
                i% = i% + 1
        Next
        Return

BTTransfer:
       Cls
        locate 4, 1
       Print LangSet$(56)
        GoSub Enable_Bluetooth
        Cls
        locate 4, 1
      Print LangSet$(57)
        a$ = ""
        While Asc(a$) <> 27
                a$ = inkey$
        Wend
        Close_Com (Comport%)
        Return


transferdata:
        Cls
        GoSub Enable_COM1
        locate 4, 1
        Print LangSet$(58)
        a$ = ""
        While Asc(a$) <> 27
                a$ = inkey$
        Wend
        Close_Com (Comport%)
        Return

Enable_IrCradle:
        set_com (ComPort%, 1, 1, 2, 1 )                               ' com, 115200 n, 8, 1
        set_com_type (ComPort%, 3)
        open_com (ComPort%)
        On com(ComPort%) GoSub net_cmd                                   ' if download from application
        Return

Enable_COM1:
        set_com ( 1, 1, 1, 2, 1 )                               ' com1, 115200 n, 8, 1
        OPEN_COM (1)
        On com(1) GoSub net_cmd                                     ' if download from application
       Comport% = 1
        Return

Enable_COM2:
      
        set_com ( 2, 1, 1, 2, 1 )                               ' com2, 115200, n, 8, 1
        set_com_type ( 2, 3 )
        OPEN_COM (2)
       Comport% = 2
        On com(2) GoSub net_cmd                             ' if download from application
        Return

Enable_Bluetooth:
        'Set com 2 as Bluetooth com port (SPP)
        SET_COM_TYPE(2,5)

        'Set_COM(N%,Baudrate%, Parity%,Data%,Handshake%)

        'N% - 2 is for SPP
        'Baudrate% - 4 is Master for SPP
        'Parity% - Not applied in Bluetooth, Set to 1
        'Data% - Not applied in Bluetooth, Set to 1
        'Handshake% - Set 4 for Wedge Emulator
        SET_COM(2,1,1,2,1)
        OPEN_COM (2)
        Comport% = 2
        On com(2) GoSub net_cmd                             ' if download from application
        Return

net_cmd:
        
        cmd$ = read_com(Comport%)
        cmd$ = Mid(cmd$, 1, Len(cmd$))
        Value$ = ""
         
       'Add Stock
       If Mid(cmd$, 1, 2) = "#S" And Len(cmd$) = 134 Then
           
            dbf% = 1
          	Data$ = Mid(cmd$, 3, Len(cmd$))         
    	GoSub addrecord

      else if Mid(cmd$, 1, 2) = "#Q" And Len(cmd$) = 35 Then
           
           dbf% = 2
           Data$ = Mid(cmd$, 3, Len(cmd$))
	GoSub addrecord
	
     else if cmd$="#PROGRAM" then

                msg$ = LangSet$(59)
                write_com(ComPort%,ProgramName$)

        else if cmd$="#TEST" then
                
                GoSub showcommstatus
                write_com(ComPort%,"#ACK")
     
       else if cmd$="#RECCOUNT" then
                
                write_com(ComPort%,str(record_count(dbf%)))
        
        else if cmd$="#CLEAR" then
                
                GoSub showcommstatus
                empty_file (dbf%)
                write_com(ComPort%,"#ACK")

        else if left(cmd$,7) ="#SETIDX" then
                
                dbf% = Val(Mid(cmd$, 8))
                GoSub showcommstatus
                write_com ( ComPort%, "#ACK" )
        
      else if cmd$ = "#GOTOP"  then

                If record_count(dbf%) > 0 Then
                        msg$ = LangSet$(60)
                        GoSub showcommstatus
                        Move_To(dbf%,1)
                        write_com ( ComPort%, Get_Record$(dbf%) )
                Else
                        msg$ = LangSet$(61)
                        GoSub showcommstatus
                        write_com ( ComPort%, "#OVER" )
                End If
        
        else  if cmd$ = "#NEXT" then

                If record_count(dbf%) = get_record_number(dbf%) Then

                        msg$ = LangSet$(62)
                        GoSub showcommstatus
                        write_com ( ComPort%, "#OVER" )

                Else
                        msg$ = LangSet$(63)
                        GoSub showcommstatus
                        Move_To_Next (dbf%)
                        write_com (ComPort%, Get_Record$(dbf%))

                End If

      else  if cmd$ = "#PROMPT" then

		write_com ( ComPort%, "#ACK" )  
		gosub prompt

     else if left(cmd$,7) ="#PREFIX" then
                
                value$ = Mid(cmd$, 8)
                if len(value$) > 0 then
		serialidprefix$ = value$
	   end if
             write_com ( ComPort%, serialidprefix$ )

   else if left(cmd$,13) ="#CUSTOMERCODE" then
                
                value$ = Mid(cmd$,14)
                if len(value$) > 0 then
		customercode$ = value$
		gosub customize
	   end if
             write_com ( ComPort%, customercode$ )

     else if left(cmd$,9) ="#SERIALID" then
                
                value$ = Mid(cmd$, 10)
                if len(value$) > 0 then
		serialid& = valr(value$)
	   end if
             write_com ( ComPort%, str(serialid&))

     else if left(cmd$,7) ="#HEADER" then
                
                value$ = Mid(cmd$, 8)
                if len(value$) > 0 then
		header$ = value$
	   end if
             write_com ( ComPort%, header$ )

    else if left(cmd$,7) ="#FOOTER" then
                
                value$ = Mid(cmd$, 8)
                if len(value$) > 0 then
		footer$ = value$
	   end if
             write_com ( ComPort%, footer$ )

      Else

                Off Timer(1)
                msg$ = LangSet$(64)
                GoSub showcommstatus
                write_com ( ComPort%, "#NAK" )

        End If

        'On Timer(1, 200) GoSub prompt
       
        Return

showcommstatus:
        If lastmsg$ <> msg$ Then
                Cls
                locate 4, 1
                Print msg$
                lastmsg$ = msg$
        End If
        Return

addrecord:
          msg$ = LangSet$(65)
          GoSub showcommstatus
          add_record ( dbf%, Data$ )
          cmd$ = ""
         unlock
         write_com ( ComPort%, "#ACK")
        Return

prompt:
        Cls
        locate 4, 1
        Print LangSet$(66)
        locate 6, 1
        Print LangSet$(67)
        beep(4000,150)
      Close_Com (Comport%)
        Off Timer(1)
        Return


statusmenu:
        GoSub InitMenu
        Menu_str$ = LangSet$(68) + Chr$(13)
        Menu_str$ = Menu_str$ + LangSet$(69) + Chr$(13)

statusmenuLoop:
                sel% = Menu(Menu_str$)
        If sel% = 1 Then
                GoSub showCashInvoicestatus
        else if sel% = 2 then
                GoSub showinventorystatus
        else if sel% = 0 then
                GoTo mainmenu
        End If
        Return statusmenuLoop

showCashInvoicestatus:
        Cls
        locate 2, 1
        Print LangSet$(70)
        
        move_to(3,1)
         locate 4, 1
        Print LangSet$(71) + Str(record_count(3))
        move_to(3,1)
        If record_count(3) > 0 Then
                EOF% = 0
                Volume! = 0
                While EOF% = 0
                        record$ = get_record(3)
                        Amt! = VALR(Mid(record$, 31, MaxLenPrice%))
                        Volume! = Volume! + Amt!
                        If record_count(3) = get_record_number(3) Then
                                EOF% = 1
                        Else
                                Move_To_Next (3)
                        End If
                Wend
        End If
        locate 6, 1
        Print LangSet$(72) + Str(Volume!)

        a$ = inkey$
        While Asc(a$) <> 27     ' Esc
                a$ = inkey$
        Wend
        return statusmenu

showinventorystatus:
        Cls
        locate 2, 1
        Print LangSet$(73)
        locate 4, 1
        Print LangSet$(74) + Str(record_count(1))
        a$ = inkey$
        While Asc(a$) <> 27     ' Esc
                a$ = inkey$
        Wend
        return statusmenu

ShowItemNotFound:

        beep ( 3000, 20, 0, 20, 3000, 20 )
        Print " Not found"
        wait (300)
        Return

showmsg:
      xloc% = 4
      yloc% = 1
       Cls                                     ' Clear Screen
       GoSub showFlash
       Return

showFlash:
       to_len% = Len(msg$)
       filldata$ = " "
       fillspacedata$ = ""
       GoSub fill_space_500
      locate xloc%, yloc%
       Print msg$
       wait (300)
      
      Return

fill_space_500:
        For c = Len(fillspacedata$) + 1 To to_len%
                fillspacedata$ = leadingdata$ + fillspacedata$ + filldata$
        Next c
       leadingdata$ = ""
        filldata$ = ""
        Return

clear1:
        locate 1, 1
      select_font(1)
            Print BlankLine$
            Return
clear2:
        locate 2, 1
      select_font(1)
        Print BlankLine$
        Return
clear3:
        locate 3, 1
      select_font(1)
        Print BlankLine$
        Return
clear4:
        locate 4, 1
      select_font(1)
        Print BlankLine$
        Return
clear5:
        locate 5, 1
      select_font(1)
        Print BlankLine$
        Return
clear6:
        locate 6, 1
      select_font(1)
        Print BlankLine$
        Return
clear7:
        locate 7, 1
      select_font(1)
        Print BlankLine$
        Return
clear8:
        locate 8, 1
      select_font(1)
        Print BlankLine$
        Return

InitMenu:
        Cls
        Off Esc
        Disable Reader(1)
        Select_font (default_font%)
        Return

SetLanguage:
                If language$ = "BIG5" Then
                        GoSub SetTradChin
                else if language$ = "GB" then
                        GoSub SetSimpChin
                Else
                        GoSub SetEnglish
                End If
                Return

        
SetEnglish:
				LangSet$(1) = "&2 Check Stock"
				LangSet$(2) = "&3 Transfer"
				LangSet$(3) = "&4 Status"
				LangSet$(4) = "&5 Printer"
				LangSet$(5) = "&6 Language"
				LangSet$(6) = "&1 English"
				LangSet$(7) = "&2 Trad. Chi"
				LangSet$(8) = "&3 Simp. Chi"
				LangSet$(9) = "&1 Enable Printer"
				LangSet$(10) = "&2 Select Printer"
				LangSet$(11) = "&4 Print Test"
				LangSet$(12) = "&5 Header"
				LangSet$(13) = "&6 Footer"
				LangSet$(14) = "&1 Enable"
				LangSet$(15) = "&2 Disable"
				LangSet$(16) = "Printer"
				LangSet$(17) = "Enabled"
				LangSet$(18) = "Printer"
				LangSet$(19) = "Disabled"
				LangSet$(20) = "&2 Citizen IR"
				LangSet$(21) = "&3 Citizen BT"
				LangSet$(22) = "&4 Custom BT"
				LangSet$(23) = "Citizen IR"
				LangSet$(24) = "selected"
				LangSet$(25) = "Citizen BT"
				LangSet$(26) = "selected"
				LangSet$(27) = "Cancelled"
				LangSet$(28) = "Custom BT"
				LangSet$(29) = "selected"
				LangSet$(30) = "Cancelled"
				LangSet$(31) = "&1 Price"
				LangSet$(32) = "&2 Quantity"
				LangSet$(33) = "Barcode:"
				LangSet$(34) = "&1 Price A"
				LangSet$(35) = "&2 Price B"
				LangSet$(36) = "&3 Price C"
				LangSet$(37) = "&4 Price D"
				LangSet$(38) = "&5 Price E"
				LangSet$(39) = "&6 Advance"
				LangSet$(40) = "Current: "
				LangSet$(41) = "New: "
				LangSet$(42) = "Barcode:"
				LangSet$(43) = "Item: "
				LangSet$(44) = "Connecting"
				LangSet$(45) = "Citizen IR"
				LangSet$(46) = "Connecting"
				LangSet$(47) = "Citizen BT"
				LangSet$(48) = "Connecting"
				LangSet$(49) = "Custom BT"
				LangSet$(50) = "&1 IR Cradle"
				LangSet$(51) = "&2 Bluetooth"
				LangSet$(52) = "  Finding   "
				LangSet$(53) = " BT Device  "
				LangSet$(54) = "Connecting "
				LangSet$(55) = "Pairing Fail"
				LangSet$(56) = "Please wait "
				LangSet$(57) = " Bluetooth     Ready    "
				LangSet$(58) = "  Ready to    transmit"
				LangSet$(59) = "Connecting.."
				LangSet$(60) = "  Uploading"
				LangSet$(61) = " Completed"
				LangSet$(62) = " Completed"
				LangSet$(63) = "  Uploading"
				LangSet$(64) = "    Error"
				LangSet$(65) = "Downloading"
				LangSet$(66) = " Completed"
				LangSet$(67) = " Press Esc"
				LangSet$(68) = "&1 Invoice"
				LangSet$(69) = "&2 Inventory"
				LangSet$(70) = "  Invoice   "
				LangSet$(71) = "Count:"
				LangSet$(72) = "Sum: "
				LangSet$(73) = " Inventory  "
				LangSet$(74) = "Item:"
				LangSet$(75) = "&1 Invoice"


				Language$ = "ENG"

				default_font% = 4

				return
				
SetTradChin:
				LangSet$(1) = "&2 琩高砯兜"
				LangSet$(2) = "&3 戈肚癳"
				LangSet$(3) = "&4 瞷戈癟"
				LangSet$(4) = "&5 诀"
				LangSet$(5) = "&6 Language"
				LangSet$(6) = "&1 English"
				LangSet$(7) = "&2 Trad. Chi"
				LangSet$(8) = "&3 Simp. Chi"
				LangSet$(9) = "&1 诀"
				LangSet$(10) = "&2 匡拒诀"
				LangSet$(11) = "&4 代刚"
				LangSet$(12) = "&5 "
				LangSet$(13) = "&6 Ю"
				LangSet$(14) = "&1 ?笆"
				LangSet$(15) = "&2 "
				LangSet$(16) = "诀"
				LangSet$(17) = "砞﹚Θ"
				LangSet$(18) = ""
				LangSet$(19) = ""
				LangSet$(20) = "&2 Citizen IR"
				LangSet$(21) = "&3 Citizen BT"
				LangSet$(22) = "&4 Custom BT"
				LangSet$(23) = "Citizen IR"
				LangSet$(24) = "砞﹚Θ"
				LangSet$(25) = "Citizen BT"
				LangSet$(26) = "砞﹚Θ"
				LangSet$(27) = ""
				LangSet$(28) = "Custom BT"
				LangSet$(29) = "砞﹚Θ"
				LangSet$(30) = ""
				LangSet$(31) = "&1 基"
				LangSet$(32) = "&2 计秖"
				LangSet$(33) = "兵絏:"
				LangSet$(34) = "&1 A基"
				LangSet$(35) = "&2 B基"
				LangSet$(36) = "&3 C基"
				LangSet$(37) = "&4 D基"
				LangSet$(38) = "&5 E基"
				LangSet$(39) = "&6 璹"
				LangSet$(40) = "瞷基: "
				LangSet$(41) = "穝基: "
				LangSet$(42) = "兵絏:"
				LangSet$(43) = "砯兜: "
				LangSet$(44) = "硈钡い"
				LangSet$(45) = "Citizen IR"
				LangSet$(46) = "硈钡い"
				LangSet$(47) = "Citizen BT"
				LangSet$(48) = "硈钡い"
				LangSet$(49) = "Custom BT"
				LangSet$(50) = "&1 絬诀畒"
				LangSet$(51) = "&2 膞"
				LangSet$(52) = "碝т"
				LangSet$(53) = "膞杆竚  "
				LangSet$(54) = "硈钡い "
				LangSet$(55) = "硈钡ア毖"
				LangSet$(56) = "叫祔"
				LangSet$(57) = "膞非称"
				LangSet$(58) = "非称肚癳"
				LangSet$(59) = "硈钡い.."
				LangSet$(60) = "  更い"
				LangSet$(61) = " ЧΘ"
				LangSet$(62) = " ЧΘ"
				LangSet$(63) = "  更い"
				LangSet$(64) = "岿粇"
				LangSet$(65) = "更い"
				LangSet$(66) = " ЧΘ"
				LangSet$(67) = "  Esc"
				LangSet$(68) = "&1 綪扳虫"
				LangSet$(69) = "&2 "
				LangSet$(70) = "  綪扳虫   "
				LangSet$(71) = "綪扳虫:"
				LangSet$(72) = "羆蝗絏: "
				LangSet$(73) = "   "
				LangSet$(74) = "砯兜:"
				LangSet$(75) = "&1 綪扳虫"
				
				Language$ = "BIG5"
				
				default_font% = 5

				return

SetSimpChin:
				LangSet$(1) = "&2 查询货项"
				LangSet$(2) = "&3 资料传送"
				LangSet$(3) = "&4 现时资讯"
				LangSet$(4) = "&5 印表机"
				LangSet$(5) = "&6 Language"
				LangSet$(6) = "&1 English"
				LangSet$(7) = "&2 Trad. Chi"
				LangSet$(8) = "&3 Simp. Chi"
				LangSet$(9) = "&1 印表机"
				LangSet$(10) = "&2 选择印表机"
				LangSet$(11) = "&4 列印测试"
				LangSet$(12) = "&5 页首"
				LangSet$(13) = "&6 页尾"
				LangSet$(14) = "&1 ?动"
				LangSet$(15) = "&2 取消列印"
				LangSet$(16) = "印表机"
				LangSet$(17) = "设定成功"
				LangSet$(18) = "取消列印"
				LangSet$(19) = ""
				LangSet$(20) = "&2 Citizen IR"
				LangSet$(21) = "&3 Citizen BT"
				LangSet$(22) = "&4 Custom BT"
				LangSet$(23) = "Citizen IR"
				LangSet$(24) = "设定成功"
				LangSet$(25) = "Citizen BT"
				LangSet$(26) = "设定成功"
				LangSet$(27) = "取消"
				LangSet$(28) = "Custom BT"
				LangSet$(29) = "设定成功"
				LangSet$(30) = "取消"
				LangSet$(31) = "&1 价格"
				LangSet$(32) = "&2 数量"
				LangSet$(33) = "条码:"
				LangSet$(34) = "&1 A价"
				LangSet$(35) = "&2 B价"
				LangSet$(36) = "&3 C价"
				LangSet$(37) = "&4 D价"
				LangSet$(38) = "&5 E价"
				LangSet$(39) = "&6 自订"
				LangSet$(40) = "现价: "
				LangSet$(41) = "新价: "
				LangSet$(42) = "条码:"
				LangSet$(43) = "货项: "
				LangSet$(44) = "连接中"
				LangSet$(45) = "Citizen IR"
				LangSet$(46) = "连接中"
				LangSet$(47) = "Citizen BT"
				LangSet$(48) = "连接中"
				LangSet$(49) = "Custom BT"
				LangSet$(50) = "&1 红外线机座"
				LangSet$(51) = "&2 篮芽"
				LangSet$(52) = "寻找"
				LangSet$(53) = "篮芽装置  "
				LangSet$(54) = "连接中 "
				LangSet$(55) = "连接失败"
				LangSet$(56) = "请稍候"
				LangSet$(57) = "篮芽准备"
				LangSet$(58) = "准备传送"
				LangSet$(59) = "连接中.."
				LangSet$(60) = "  上载中"
				LangSet$(61) = " 完成"
				LangSet$(62) = " 完成"
				LangSet$(63) = "  上载中"
				LangSet$(64) = "错误"
				LangSet$(65) = "下载中"
				LangSet$(66) = " 完成"
				LangSet$(67) = " 按 Esc"
				LangSet$(68) = "&1 销售单"
				LangSet$(69) = "&2 仓存"
				LangSet$(70) = "  销售单   "
				LangSet$(71) = "销售单:"
				LangSet$(72) = "总银码: "
				LangSet$(73) = " 仓存  "
				LangSet$(74) = "货项:"
				LangSet$(75) = "&1 销售单"

				language$ = "GB"

				default_font% = 5

				return

SetPrinterLanguage:

                If Printerlanguage$ = "BIG5" Then
                        GoSub SetTradChinPrinter
                else if Printerlanguage$ = "GB" then
                        GoSub SetSimpChinPrinter
                Else
                        GoSub SetEnglishPrinter
                End If
                Return


SetEnglishPrinter:
			printerset$(1) = "   Discount"
			printerset$(2) ="   Total"
			printerset$(3)= "   Paid"
			printerset$(4)= "   Change"
			printerset$(5) = "Cash Invoice"
			printerset$(6) = "   Sub total"
			PrinterSet$(7) = "                  VER. "
			PrinterSet$(8) = "                 PRINT TEST                 "
			PrinterSet$(9) = "                   HEADER                   "
			PrinterSet$(10) = "NULL"
			PrinterSet$(11) = "                   FOOTER                   "
			PrinterSet$(12) = "NULL"
			PrinterSet$(13) = "            DOWNLOADABLE GRAPHIC            "
			PrinterSet$(14) = "                    END                     "
			
			printerlanguage$ = "ENG"
		
			return


SetTradChinPrinter:
			
			Printerset$(1) = "   чΙ"
			Printerset$(2) ="   羆计"
			Printerset$(3)= "   蹿"
			Printerset$(4)= "   т尿"
			Printerset$(5) = "綪扳虫"
			Printerset$(6) = "   璸"
			PrinterSet$(7) = "                  セ "
			PrinterSet$(8) = "                 代刚                 "
			PrinterSet$(9) = "                                      "
			PrinterSet$(10) = "礚"
			PrinterSet$(11) = "                   Ю                   "
			PrinterSet$(12) = "礚"
			PrinterSet$(13) = "                 ず竚瓜                 "
			PrinterSet$(14) = "                    Ч                     "
			
			printerlanguage$ = "BIG5"
		
			return

SetSimpChinPrinter:
			
			printerset$(1) = "   Discount"
			printerset$(2) ="   Total"
			printerset$(3)= "   Paid"
			printerset$(4)= "   Change"
			printerset$(5) = "Cash Invoice"
			printerset$(6) = "   Sub total"
			PrinterSet$(7) = "                  VER. "
			PrinterSet$(8) = "                 PRINT TEST                 "
			PrinterSet$(9) = "                   HEADER                   "
			PrinterSet$(10) = "NULL"
			PrinterSet$(11) = "                   FOOTER                   "
			PrinterSet$(12) = "NULL"
			PrinterSet$(13) = "            DOWNLOADABLE GRAPHIC            "
			PrinterSet$(14) = "                    END                     "
			
			printerlanguage$ = "GB"
			return

GetTextFont:
			fonttype% = 1
			MaxChar% = MaxLineChar% 
			for idx% = 1 to len(text$)
				c$  = mid(text$,idx%,1)					
				if asc(c$) > 127 or asc(c$) < 33 then
					fonttype% = 2
					MaxChar% = 6
				end if				
			next
			return

SetupMode:
	gosub TransferMenu
	return


ReadBarcode:
	data$ = get_reader_data$(1)
	gosub customizeinput
	return

customizeinput:
	if mod1% = 1 then
		if len(data$) = 13 then
			if mid(data$,1,3) = "977" or mid(data$,1,3) = "978" or mid(data$,1,3) = "979" then
				data$ = mid(data$,4)
				data$ = left(data$,len(data$)-1)
			end if
		end if 
	end if
	return

EnterCustomerCode:
	cls
	locate 2,1
	print "Cust. Code"
	locate 4,1
	cursor(1)
	input customercode$
	cursor(0)
	gosub customize
	return

Customize:

	if ucase(customercode$) = "COSMOS" then
		gosub enablecosmos
	end if
	return

EnableCOSMOS:
	mod1% = 1
	return