#!/bin/sh

# !!! jq must be installed !!!

key=b2e9acf7-09cb-4e72-8d66-588cbcfd8a30



file1=`date "+%Y-%m-%d-%H-%M-%S"`



new_url=`echo https://support.russianpost.ru/sd/services/rest/create-m2m/report%24rep5167/%7BcheckInTime:%222016.12.01%2000:00%22,periodDate:%222016.12.31%2023:59%22%7D?accessKey=${key}`
curl $new_url >${file1}_rep_1 2>>${file1}_rep_11
#todo add err checking

id_rep=`cat ${file1}_rep_1|jq '[.UUID]'|sed "s/\"//g"|cut -s -d "$" -f 2`

echo Report ID=$id_rep please wait up to 60 seconds...
w=0
while  [ $w -eq 0 ];
do
new_url=`echo https://support.russianpost.ru/sd/services/rest/get/report%24${id_rep}?accessKey=${key}`
curl $new_url >${file1}_rep_2 2>>${file1}_rep_22
#todo add err checking

f_id=`cat ${file1}_rep_2|jq '.files[]|.UUID'|sed "s/\"//g"|cut -d "$" -f 2 `
f_name=`cat ${file1}_rep_2|jq '.files[]|.title'`

if [ -n "$f_id" ]; then
    echo "ready"
w=1
else
    echo "wait..."
w=0
fi

sleep 5
rm -f ${file1}_rep_1
rm -f ${file1}_rep_2
rm -f ${file1}_rep_11
rm -f ${file1}_rep_22

done

echo $f_id
echo $f_name
new_url=`echo https://support.russianpost.ru/sd/services/rest/get-file/file%24$f_id?accessKey=${key}`
curl -o result.xlsx $new_url

exit

