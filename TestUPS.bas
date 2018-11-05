'To test UPS / STAR TSP743
'2005.5.23

interval% = 3
printcount% = 0
Set_Com_Type (1, 3)
set_com (1, 5, 1, 2, 1 )                               ' com, 19200, n, 8, 1
IRDA_TIMEOUT (1)
open_com(1)
locate 1,1
print "Interval (sec):"
locate 2,1
input interval%
on timer(1,interval% * 100) gosub prn
locate 3,1
print "Start time:"
locate 4,1
print Date$ + " " + Time$

input a$

prn:
	idx%  =0
	printcount% = printcount% + 1
	data$ = Date$ + " " + Time$ + " - " + str(printcount%) + chr(10)
	gosub sendbuffer
	while (idx% < 51)
		
		idx% = idx% + 1
		locate 1,1
		data$ = "A就會如雪候就會如雪花飛至就會如雪花飛至" + chr(10)
		gosub sendbuffer

	wend
	
	locate 7,1
	print "Last print: " 
	locate 8,1
	print date$ + " " + time$
	return

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
	Write_Com(1, senddata$)
	return
