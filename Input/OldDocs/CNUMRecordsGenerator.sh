##############################################################################
# Created By: Amit(Aricent)
# Description: Shell Script to generate the CNUM Bulk-Provisioning Input file
# Usage ./shellscript iccidstart imsistart records_to_be_generated  
##############################################################################

if [ "$#" -ne 3 ]
then
	echo "USAGE: ./script_name ICCID_START IMSI_START nmbrOfRecords" 
	echo "ERROR: Arguments not correct, Exiting........."
else
	if [ -f ./BulkSimProvisioningList.txt ]
	then
		rm -rf ./BulkSimProvisioningList.txt
	fi

	if [ -f ./Bulk_Provisioning_OrderCare.txt ]
        then
                rm -rf ./Bulk_Provisioning_OrderCare.txt
        fi

			
	ICCID=$1
	IMSI=$2
	iterator=$3
	MSN=1001
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

if [ $iterator -ge 1000 ]

then

	echo "ICCID	IMSI	PIN1	PUK1	PIN2	PUK2	KI	ENCRYPTED	KI	OPC	ENCRYPTED	OPC	TRANSKEY	TRANSKEY	INDEX	MSN	IMPU	IMPI" >> BulkSimProvisioningList.txt

fi
	for ((i=0; i<$iterator;i++))
	do	
		if [ $batchCheck -eq  500 ]
		then
			MSNtmp=`expr $MSNtmp + 1`
			batchCheck=0
		fi
		ICCIDtmp=`expr $1 + $i`
		IMSItmp=`expr $2 + $i`
		echo "$ICCIDtmp""9	$IMSItmp	1234	12345678	1234	12345678	4FF3F7654BCF7A4A2A12A5FB7466F888	4FF3F7654BCF7A4A2A12A5FB7466F888	E3B4D1987C0E577A09FF75AEC158F999	E3B4D1987C0E577A09FF75AEC158F999	0	0	$MSNtmp	SIP:""$IMSItmp""@ims.mnc874.mcc405.3gppnetwork.org	$IMSItmp""@ims.mnc874.mcc405.3gppnetwork.org" >> BulkSimProvisioningList.txt
		batchCheck=`expr $batchCheck + 1`
	done
		


#########LOGIC FOR CREATING ORDER CARE INPUT FILE########################################

		mod=`expr $iterator % 500`
		echo "MSN|StartICCIDRange|EndICCIDRange|Flag" >> Bulk_Provisioning_OrderCare.txt
		for ((i=1; i<=$orderCareRecords;i++))
		do
			if [ $i -eq $orderCareRecords ] && [ $mod -ne 0 ]
			then
				#mod=`expr $iterator % 500`
				 	
				ICCIDEnd=`expr $ICCIDStart  + $mod - 1`
				echo "$MSNtmp1|$ICCIDStart|$ICCIDEnd|M" >> Bulk_Provisioning_OrderCare.txt
			else
				  
				ICCIDEnd=`expr $ICCIDStart  + 499`
				echo "$MSNtmp1|$ICCIDStart|$ICCIDEnd|M" >> Bulk_Provisioning_OrderCare.txt
				ICCIDStart=`expr $ICCIDEnd + 1`
				#echo "$MSNtmp1|$ICCIDStart|$ICCIDEnd|M"
			fi
			#echo "$MSNtmp1|$ICCIDStart|$ICCIDEnd|M"
		MSNtmp1=`expr $MSNtmp1 + 1`
		done
	echo "File Creation Successful" 
fi

