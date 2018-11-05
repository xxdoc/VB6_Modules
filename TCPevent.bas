		SET_COM(1,1,1,2,1)
		SET_COM_TYPE(1,3)
        OPEN_COM(1)
		'START_DEBUG(1,1,1,2,1)
       ' GOSUB SET_TCPIP
		A$ = "0123456789"
		B$ = ""

		FOR I= 1 TO 10 
			B$=B$+A$
		NEXT I
		LOCATE 8,1  
		PRINT "I AM :" ; GET_TARGET_MACHINE

		START TCPIP
		ON TCPIP GOSUB TCPIP_TRIGGER

 
		NOW% = TIMER
LOOP:   '
        '**********************************
        'WRITE YOUR PORGRAM HERE
        '**********************************
        '
		IF TIMER <> NOW% THEN
		LOCATE 1,1
		PRINT "                    "
		LOCATE 1,1  
		PRINT "C:" ; GET_WLAN_STATUS(1)
		LOCATE 1,6
		PRINT "Q:" ; GET_WLAN_STATUS(2)
		LOCATE 1,11 
		PRINT "S:" ; GET_WLAN_STATUS(3)
		LOCATE 1,16
		PRINT "N:" ; GET_WLAN_STATUS(4)
		LOCATE 2,1
		PRINT "===================="
		NOW% = TIMER
		END IF
        GOTO LOOP

        STOP TCPIP

'
'
'********************************************************************************
'                TCPIP TRIGGER FUNCTION EXAMPLE
'       Please note the value of variable will be changed in event function
'********************************************************************************
'
TCPIP_TRIGGER:
        '
        '*************************************************************************************
        ' "GET_TCPIP_MESSAGE" MUST BE THE FIRST COMMAND IN EVENT TRIGGER
        '*************************************************************************************
        MSG=GET_TCPIP_MESSAGE
        '
        '
        '********************************
        'IP READY TRIGGER
        '********************************
        '
        IF MSG >= 4080 THEN
        	GOSUB GET_TCPIP_SETTING
			BEEP(1100,10)
	        '
	        '*************************************
	        'OPEN TWO CONNECTIONS AS SERVER MODE
	        '*************************************
	        '
		TCP_OPEN(0,"0.0.0.0",0,23,0,13)
		TCP_OPEN(1,"0.0.0.0",0,24,0,12)
		TCP_OPEN(2,"192.168.2.140",1024,0,0,11)

        '
        '**************************
        'DATA CHECK
        '**************************
        '
        ELSE IF MSG >= 4060 THEN
            CNT=MSG-4060
			A$=NREAD(CNT)
			IF B$ = A$ THEN
				BEEP(4400,4)
			ELSE
				BEEP(8800,4)
			END IF	
        	NWRITE(CNT,B$)
        '
        '**************************
        'BREAK EVENT CHECK
        '**************************
        '
        ELSE IF MSG >= 4040 THEN
			BEEP(2200,4)
        '
        '*****************************
        'CONNECT EVENT CHECK
        '*****************************
        '
        ELSE IF MSG >= 4020 THEN
            CNT=MSG-4020
			NWRITE(CNT,B$)
        '
        '********************************
        'BUFFER OVER FLOW
        '********************************
        '
        ELSE IF MSG >= 4000 THEN
			BEEP(1100,10)
        END IF
        '
        '**********************************
        'WRITE YOUR PORGRAM HERE
        '**********************************
        '           
       RETURN
        
GET_TCPIP_SETTING:
        '
        '******************************
        'GET TCPIP SETTING
        '******************************
        '
        LOCATE 3,1
        IP$ = SOCKET_IP(-1)
        PRINT "I:",IP$
        
        IP$ = SOCKET_IP(-2)
        PRINT "M:",IP$
        
        IP$ = SOCKET_IP(-3)
        PRINT "R:",IP$
        
        IP$ = SOCKET_IP(-4)
        PRINT "D:",IP$
        RETURN

SET_TCPIP:
         '
        '*****************************************************************
        'SET LOCAL IP,SUBNET MASK,ROUTER,DNS SERVER
        'ALL OF THE VALUE WILL BE WRITE TO FLASH
        '      !! DO NOT  TO CONFIG TCPIP EVERY TIME !!
        '*****************************************************************
        '
        'SET LOCAL IP
        IP_CFG(1,"210.242.202.248")
        'SET SUBNET MASK
        IP_CFG(2,"255.255.255.0")
        'SET ROUTER
        IP_CFG(3,"210.242.202.254")
        'SET DNS SERVER
        IP_CFG(4,"168.95.1.1")
        'SET DHCP ENABLE
        IP_CFG(11,"Enable")
        'SET Autnen ENABLE
        IP_CFG(12,"Enable")
        'SET DEFAULT WEP KEY
        IP_CFG(15,"2")

        'SET WEP KEY2 = ROOT
        IP_CFG(8,"ROOT")


        RETURN
