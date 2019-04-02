##############################################################################
# Created By: Amit(Aricent)
# Description: Shell Script to generate the CNUM Bulk-Provisioning Input file
# Usage ./shellscript iccidstart imsistart records_to_be_generated  
##############################################################################

if [ "$#" -ne 5 ]
then
	echo "USAGE: ./script_name ICCID_START IMSI_START nmbrOfRecords PONumber Circle" 
	echo "ERROR: Arguments not correct, Exiting........."
else
	if [ -f ./SIMProvisioning.txt ]
	then
		rm -rf ./SIMProvisioning.txt
	fi

	if [ -f ./SIM_PROV.txt ]
        then
                rm -rf ./SIM_PROV.txt
        fi

			
	ICCID=$1
	IMSI=$2
	iterator=$3
	ponmbr=$4
	Circle=$5
	MSN=1601
    batchCheck=0
	MSNtmp=$MSN
	ICCIDtmp=$ICCID
	IMSItmp=$IMSI
	ICCIDStart=$1
	MSNtmp1=$MSN


		
	orderCareRecords=`expr $iterator / 500`
	
	if [ `expr $iterator % 500` -ne 0 ]
	then
		orderCareRecords=`expr $orderCareRecords + 1`
	fi	 
#	echo $orderCareRecords

############################Header CREATION FOR CNUM BULK FILE###########################################

##if [ $iterator -ge 1000 ]

##then

##	#echo "ICCID	IMSI	PIN1	PUK1	PIN2	PUK2	KI	ENCRYPTED	KI	OPC	ENCRYPTED	OPC	TRANSKEY	TRANSKEY	INDEX	MSN	IMPU	IMPI" >> SIMProvisioning.txt
##	echo "IMPU	IMPI	IMSI	ICCID	PIN1	PUK1	PIN2	PUK2	EncryptedKi	EncryptedOPC	TranskeyIndex	OPKeyIndex	MSN" >> SIMProvisioning.txt

	echo "PO Number: $ponmbr" >> SIMProvisioning.txt
	echo "Batch NO: $ponmbr" >> SIMProvisioning.txt
	echo "SIM Type: ISIM" >> SIMProvisioning.txt
	echo "Circle: $Circle" >> SIMProvisioning.txt
	echo "SIM Quantity: $iterator" >> SIMProvisioning.txt
	echo "SIM Subtype: RAW" >> SIMProvisioning.txt
	echo "Memory: 128K" >> SIMProvisioning.txt
	echo "SIM Size: 2FF+3FF+4FF" >> SIMProvisioning.txt
	echo "Card Type: SIM" >> SIMProvisioning.txt
	echo "SKU: 1108" >> SIMProvisioning.txt
	echo "Card Manufacturer: 110429" >> SIMProvisioning.txt
	echo "*************************************************" >> SIMProvisioning.txt
	echo "* INPUT AND OUTPUT VARIABLES DESCRIPTION" >> SIMProvisioning.txt
	echo "*************************************************" >> SIMProvisioning.txt
	echo "IMPU	IMPI	IMSI	ICCID	PIN1	PUK1	PIN2	PUK2	EncryptedKi	EncryptedOPC	TranskeyIndex	OPKeyIndex	MSN" >> SIMProvisioning.txt

##fi
	for ((i=0; i<$iterator;i++))
	do	
		if [ $batchCheck -eq  500 ]
		then
			MSNtmp=`expr $MSNtmp + 1`
			batchCheck=0
		fi
		ICCIDtmp=`expr $1 + $i`
		IMSItmp=`expr $2 + $i`
	#			ICICD		IMSI		PIN1	PUK1		PIN2	PUK2		KI	ENCRYPTED						KI									OPC	ENCRYPTED						OPC									TRANSKEYindex	OPKEY	MSN	IMPU								IMPI
	#	echo "$ICCIDtmp""9	$IMSItmp	1234	12345678	1234	12345678	11112222333344445555666677778888	11112222333344445555666677778888	E3B4D1987C0E577A09FF75AEC158F999	E3B4D1987C0E577A09FF75AEC158F999	997	96	$MSNtmp	SIP:""$IMSItmp""@ims.mnc874.mcc405.3gppnetwork.org	$IMSItmp""@ims.mnc874.mcc405.3gppnetwork.org" >> SIMProvisioning.txt
		echo "SIP:$IMSItmp@ims.mnc874.mcc405.3gppnetwork.org	$IMSItmp@ims.mnc874.mcc405.3gppnetwork.org	$IMSItmp	$ICCIDtmp	1234	12345678	1234	12345678	11112222333344445555666677778888	E3B4D1987C0E577A09FF75AEC158F999	997	96	$MSNtmp" >> SIMProvisioning.txt

		batchCheck=`expr $batchCheck + 1`
	done
	#Convert to dos format
	unix2dos SIMProvisioning.txt
		


#########LOGIC FOR CREATING ORDER CARE INPUT FILE########################################

	#	mod=`expr $iterator % 500`
	#	echo "MSN|StartICCIDRange|EndICCIDRange|ServiceTypeFlag" >> SIM_PROV.txt
	#	for ((i=1; i<=$orderCareRecords;i++))
	#	do
	#		if [ $i -eq $orderCareRecords ] && [ $mod -ne 0 ]
	#		then
	#			#mod=`expr $iterator % 500`
				 	
	#			ICCIDEnd=`expr $ICCIDStart  + $mod - 1`
	#			echo "$MSNtmp1|$ICCIDStart|$ICCIDEnd|V" >> SIM_PROV.txt
	#		else
				  
	#			ICCIDEnd=`expr $ICCIDStart  + 499`
	#			echo "$MSNtmp1|$ICCIDStart|$ICCIDEnd|V" >> SIM_PROV.txt
	#			ICCIDStart=`expr $ICCIDEnd + 1`
	#			#echo "$MSNtmp1|$ICCIDStart|$ICCIDEnd|V"
	#		fi
	#		#echo "$MSNtmp1|$ICCIDStart|$ICCIDEnd|V"
	#	MSNtmp1=`expr $MSNtmp1 + 1`
	#	done
	echo "File Creation Successful" 
fi

