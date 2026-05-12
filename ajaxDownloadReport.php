<?php
session_start();
include_once('PHP_XLSXWriter/xlsxwriter.class.php');
  ini_set('display_errors', 0);
  ini_set('log_errors', 1);
  error_reporting(E_ALL & ~E_NOTICE);

    // $hostname = '192.168.13.101';
    // $username = 'phpmyadmin';
    // $password = 'b@dcr3dit';
    // $database = 'aaca_backup';
    
    // $conn = mysqli_connect($hostname,$username,$password,$database);

include_once('../config.php');
$accounts = implode(",",$_POST['myCheckboxes']);
if($_SESSION['userType']==1){
  $clientlogincode='';
  $RMSBRGLVL2clientcod='';
}else{
   $clientlogincode= "and A.ORGCODE=". "'".$_SESSION['clientCode']."' ";
   $RMSBRGLVL2clientcod="and B.RMSBRGLVL2=". "'".$_SESSION['clientCode']."' ";
  }
  if(isset($_POST['client'])){

  if($accounts == 'open' && $_POST['outputFormat'] == '2')
 {

  if($_POST['userType'] == 3)
  {
    $where = "A.ORGCODE = '".$_POST['client']."'";
  }
  else if ($_POST['userType'] == 2 || $_POST['userType'] == 4)
  {
    $where = "A.CURR_ATTY_CD = '".$_POST['client']."' ";
  }

    // $dir = '/var/www/html/bi/dist/Report/';

    $filename = "openInventoryReport.xlsx";

    header('Content-disposition: attachment; filename="'.XLSXWriter::sanitize_filename($filename).'"');
    header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    header('Content-Transfer-Encoding: binary');
    header('Cache-Control: must-revalidate');
    header('Pragma: public'); 
   
    

    


    $query= "SELECT A.WFNAME AS 'Name', A.WFFIRMFILE AS 'Firm File Number', A.ACCT_NUM AS 'Account Number', A.MSTPOENAM1 AS 'Employer Name',
(CASE WHEN A.WFPIPEB = 1 THEN 'BLUE'
      WHEN A.WFPIPEO = 1 THEN 'ORANGE'
      WHEN A.WFPIPEG = 1 THEN 'GREEN'
      WHEN A.WFPIPER = 1 THEN 'RED'
      WHEN A.WFPIPEX = 1 THEN 'PURPLE' ELSE 'WHITE' END) AS 'Pipe Status', A.CURR_STS_CD AS 'Cur Sts Code', A.CURR_STS_DESC AS 'Current Status Description',
(CASE WHEN A.WFFSTSDTF = '0001-01-01' THEN NULL ELSE A.WFFSTSDTF END) AS 'Date Last Status Changed by Firm',
(CASE WHEN A.WFFSTSDTF = '0001-01-01' THEN NULL ELSE A.WFFSTSDTF END) AS 'Date Last Status Received on Pipeway', A.LST_NTE_CD AS 'Last Note Code',
A.NTE_DESC AS 'Last Note Code Description',
(CASE WHEN A.LST_NTE_DT = '0001-01-01' THEN NULL ELSE A.LST_NTE_DT END) AS 'Last Note Date', A.LST_NTE_CD AS 'Last Activity Code', A.NTE_DESC AS 'Last Activity Description',
(CASE WHEN A.LST_ACT_DTE = '0001-01-01' THEN NULL ELSE A.LST_ACT_DTE END) AS 'Last Activity Date', A.WFDATEA AS 'Original Plmt Date',
(CASE WHEN A.WFDATEF = '0001-01-01' THEN NULL ELSE A.WFDATEF END) AS 'Date Placed at Firm', A.ORG_PL_AMT AS 'Original Plmt Amount',
A.ORGCODE AS 'Client Code', CONCAT_WS(' ', D.RCLNM1, D.RCLNM2) AS 'Client Name', A.CLT_CDE AS 'Port Code', A.CLNT_DESC AS 'Portfolio Description',
A.CURR_ATTY_CD AS 'Current Firm Code',
A.CURR_ATTY_NME AS 'Current Firm Name', A.DBT_ST_RES AS 'ST', (CASE WHEN A.LST_PYMT_DT_A = '0001-01-01' THEN NULL ELSE A.LST_PYMT_DT_A END) AS 'Last Pay Date',
A.PYMT_TRAN_CD AS 'Last Pay Trans Code', B.SYSDESC AS 'Last Pay Trans Code Desc', A.LST_PYMT_AMT AS 'Last Amount Recovered', A.TTL_PD_ACCT AS 'Total Amount Recovered',
CAST(A.DTE_LST_CC AS DATE) AS 'Last Court Cost Date',
A.TTL_CC AS 'Total Court Cost', (CASE WHEN A.JGMT_DTE = '0001-01-01' THEN NULL ELSE A.JGMT_DTE END) AS 'Judg Date', A.JGMT_AMT AS 'Judg Amt',
A.JGMT_INT_RTE AS 'Judg Int Rate', A.JGMT_CSE_NBR AS 'Case Number', A.CRT_NME AS 'Court', CAST(C.PROCDT AS DATE) 'Date Last Balance Received From Firm',
C.HTOTBAL AS 'Last Balance Received From Firm',
(CASE WHEN  A.CURR_STS_CD IN (93,994) THEN 'SIF'
      WHEN A.CURR_STS_CD = 995 THEN 'PIF' ELSE NULL END) AS 'SIF/PIF',
A.HEMPINFO AS 'POE Flag', E.TRANDTE AS 'Date POE Updated',
(CASE WHEN A.HAJDMT = 1 THEN 'Judgment'
      WHEN A.HSERV = 1 THEN 'Service'
      WHEN A.HSUIT = 1 THEN 'Suit Filed'
      WHEN A.H1DLTR = 1 THEN '1st Demand Letter' ELSE 'None' END) AS 'Current Mile Marker',
A.WFIMGPLFG AS 'Placement Documents', A.WFIMGJDFG AS 'Judgment Copy',
(CASE WHEN A.HAJDMT = 1 THEN WFJDEPA
      WHEN A.HSERV = 1 THEN WFSERVEA
      WHEN A.HSUIT = 1 THEN WFSUITEA
      WHEN A.H1DLTR = 1 THEN WF1DLTREA ELSE 0 END)  AS 'Days to Current Mile Marker',
(CASE WHEN A.HAJDMT = 1 THEN DATEDIFF(CURRENT_DATE(), JGMT_DTE)
      WHEN A.HSERV = 1 THEN DATEDIFF(CURRENT_DATE(), WFSERVDT)
      WHEN A.HSUIT = 1 THEN DATEDIFF(CURRENT_DATE(), WFSUITDT)
      WHEN A.H1DLTR = 1 THEN DATEDIFF(CURRENT_DATE(), WF1DLTRDT) ELSE 0 END) AS 'Days Since Current Mile Marker',
A.H1DLTR AS '1DL Sent Flag',
(CASE WHEN A.WF1DLTRDT <> '0001-01-01' THEN A.WF1DLTRDT ELSE NULL END) AS 'Date 1DL Sent',
A.WF1DLTREA AS 'Days to 1DL Sent',
(CASE WHEN A.WF1DLTRDT <> '0001-01-01' THEN DATEDIFF(CURRENT_DATE(), A.WF1DLTRDT) ELSE NULL END) AS 'Days Since 1DL Sent',
A.HSUIT AS 'Suit Filed Flag',
(CASE WHEN A.WFSUITDT <> '0001-01-01' THEN A.WFSUITDT ELSE NULL END) AS 'Reported Date Suit Filed',
A.WFSUITEA AS 'Days to Suit Filed',
(CASE WHEN A.WFSUITDT <> '0001-01-01' THEN DATEDIFF(CURRENT_DATE(), A.WFSUITDT) ELSE NULL END) AS 'Days Since 1st Filed',
A.HSERV AS 'Service Flag',
(CASE WHEN A.WFSERVDT <> '0001-01-01' THEN A.WFSERVDT ELSE NULL END) AS 'Date of Service',
A.WFSERVEA AS 'Days to Service',
(CASE WHEN A.WFSERVDT <> '0001-01-01' THEN DATEDIFF(CURRENT_DATE(), A.WFSERVDT) ELSE NULL END) AS 'Days Since Service',
A.HAJDMT AS 'Judg Flag',
(CASE WHEN A.JGMT_DTE <> '0001-01-01' THEN A.JGMT_DTE ELSE NULL END) AS 'Date of Judgmt',
A.WFJDEPA AS 'Days to Judgmt',
(CASE WHEN A.JGMT_DTE <> '0001-01-01' THEN DATEDIFF(CURRENT_DATE(), A.JGMT_DTE) ELSE NULL END) AS 'Days Since Judgmt',
A.HACL AS 'Closed Flag',
(CASE WHEN A.WFLCLDT <> '0001-01-01' THEN A.WFLCLDT ELSE NULL END) AS 'Date Closed',
A.WFLCLDYA AS 'Days to Closed',
(CASE WHEN A.WFLCLDT <> '0001-01-01' THEN DATEDIFF(CURRENT_DATE(), A.WFLCLDT) ELSE NULL END) AS 'Days Since Closed',
A.HALIEN AS 'Lien Flag',
(CASE WHEN A.WFLIENDT = '0001-01-01' THEN NULL ELSE A.WFLIENDT END) AS 'Date Lien',
/* AH06052026 - Handle invalid charge off dates to prevent Excel invalid date conversion */
CASE
    WHEN A.RMSDTCHGDT IN ('00000000','0','00010101')
         OR A.RMSDTCHGDT IS NULL
         OR A.RMSDTCHGDT = ''
    THEN NULL
    ELSE STR_TO_DATE(A.RMSDTCHGDT, '%Y%m%d')
END AS RMSDTCHGDT, 
A.WFORGPRBAL
FROM HSFLCLNTWF A
LEFT JOIN RMSPSYSCDE B ON A.PYMT_TRAN_CD = B.SYSCODETYP AND B.RMSRECTYPE = 1
LEFT JOIN HUBRELBAGD C ON A.ACCT_NUM = C.RACCTNUM
LEFT JOIN RMRMCLNM_REQ_DATA D ON A.ORGCODE = D.RMSBRGLVL2
LEFT JOIN (
           SELECT FILENUM, MAX(CAST(TRANDTE AS DATE)) AS TRANDTE
           FROM HSTEMPCHG
           WHERE FLDCODE = 'MASENM'
           GROUP BY FILENUM
           ) E ON A.WFFILNUM = E.FILENUM
WHERE ".$where."
AND A.CURR_STS_CD NOT LIKE '9%'";//echo  $query;exit;


    $result = mysqli_query($conn,$query);
    //print_r($result);exit;

    $writer = new XLSXWriter();


$header = array(
			    "Name" => "string",
				"Firm File Number" => "string",
				"Account Number" => "string",
				"Employer Name" => "string",
				"Pipe Status" => "string",
				"Cur Sts Code" => "string",
				"Current Status Description" => "string",
				"Date Last Status Changed by Firm" => "MM-DD-YYYY",
				"Date Last Status Received on Pipeway" => "MM-DD-YYYY",
				"Last Note Code" => "string",
				"Last Note Code Description" => "string",
				"Last Note Date" => "MM-DD-YYYY",
				"Last Activity Code" => "string",
				"Last Activity Description" => "string",
				"Last Activity Date" => "MM-DD-YYYY",
				"Original Plmt Date" => "MM-DD-YYYY",
				"Date Placed at Firm" => "MM-DD-YYYY",
				"Original Plmt Amount" => "dollar",
				"Client Code" => "string",
				"Client Name" => "string",
				"Port Code" => "string",
				"Portfolio Description" => "string",
				"Current Firm Code" => "string",
				"Current Firm Name" => "string",
				"ST" => "string",
				"Last Pay Date" => "MM-DD-YYYY",
				"Last Pay Trans Code" => "string",
				"Last Pay Trans Code Desc" => "string",
				"Last Amount Recovered" => "dollar",
				"Total Amount Recovered" => "dollar",
				"Last Court Cost Date" => "MM-DD-YYYY",
				"Total Court Cost" => "dollar",
				"Judg Date" => "MM-DD-YYYY",
				"Judg Amt" => "dollar",
				"Judg Int Rate" => "dollar",
				"Case Number" => "string",
				"Court" => "string",
				"Date Last Balance Received From Firm" => "MM-DD-YYYY",
				"Last Balance Received From Firm" => "string",
				"SIF/PIF" => "string",
				"POE Flag" => "integer",
				"Date POE Updated" => "MM-DD-YYYY",
				"Current Mile Marker" => "string",
				"Placement Documents" => "string",
				"Judgment Copy" => "string",
				"Days to Current Mile Marker" => "integer",
				"Days Since Current Mile Marker" => "integer",
				"1DL Sent Flag" => "integer",
				"Date 1DL Sent" => "MM-DD-YYYY",
				"Days to 1DL Sent" => "integer",
				"Days Since 1DL Sent" => "integer",
				"Suit Filed Flag" => "integer",
				"Reported Date Suit Filed" => "MM-DD-YYYY",
				"Days to Suit Filed" => "integer",
				"Days Since 1st Filed" => "integer",
				"Service Flag" => "integer",
				"Date of Service" => "MM-DD-YYYY",
				"Days to Service" => "integer",
				"Days Since Service" => "integer",
				"Judg Flag" => "integer",
				"Date of Judgmt" => "MM-DD-YYYY",
				"Days to Judgmt" => "integer",
				"Days Since Judgmt" => "integer",
				"Closed Flag" => "integer",
				"Date Closed" => "MM-DD-YYYY",
				"Days to Closed" => "integer",
				"Days Since Closed" => "integer",
				"Lien Flag" => "integer",
				"Date Lien" => "MM-DD-YYYY",
				"Charge Off Date" => "MM-DD-YYYY",
				"Charge Off Amount" => "dollar",
				);


		      $style = array(
	                'font-style' => 'bold', 'fill'=>'#eee', 'halign'=>'center', 'border'=>'left,right,top,bottom', 'widths'=>[25,25,25,25,25,25,40,40,40,25,40,25,25,36,25,25,25,25,25,25,25,40,25,40,15,25,25,30,25,40,25,25,25,25,25,25,25,40,40,25,25,25,25,25,25,30,30,35,35,25,25,25,25,25,25,25,25,25,25,25,25,25,25,25,25,25,25,25,25,25,25]
	              );
		    
		    
		    $writer->writeSheetHeader('OpenAccounts', $header, $style);
		   
		    while ($row = mysqli_fetch_assoc($result))
		    {
		        
		        $writer->writeSheetRow('OpenAccounts', $row);
		    }

		  

		    
		    $writer->writeToFile(str_replace(__FILE__,'/var/www/html/bi/dist/Report/uploadsexcel/'.$filename.'',__FILE__));

		   echo json_encode(array('status' => 1));
		    //echo $writer->writeToString();
		    exit(0);

	// 	    $filename = $dir.$filename;

	// 	  return json_encode($filename);

	// exit();	    

  }

else if($accounts == 'close' && $_POST['outputFormat'] == '2')
 {
    if($_POST['userType'] == 3)
  {
    $where = "A.ORGCODE = '".$_POST['client']."'";
  }
  else if ($_POST['userType'] == 2 || $_POST['userType'] == 4)
  {
    $where = "A.CURR_ATTY_CD = '".$_POST['client']."' ";
  }


   $filename = "openInventoryReport.xlsx";

    header('Content-disposition: attachment; filename="'.XLSXWriter::sanitize_filename($filename).'"');
    header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    header('Content-Transfer-Encoding: binary');
    header('Cache-Control: must-revalidate');
    header('Pragma: public'); 


     $query= "SELECT A.WFNAME AS 'Name', A.WFFIRMFILE AS 'Firm File Number', A.ACCT_NUM AS 'Account Number', A.MSTPOENAM1 AS 'Employer Name',
(CASE WHEN A.WFPIPEB = 1 THEN 'BLUE'
      WHEN A.WFPIPEO = 1 THEN 'ORANGE'
      WHEN A.WFPIPEG = 1 THEN 'GREEN'
      WHEN A.WFPIPER = 1 THEN 'RED'
      WHEN A.WFPIPEX = 1 THEN 'PURPLE' ELSE 'WHITE' END) AS 'Pipe Status', A.CURR_STS_CD AS 'Cur Sts Code', A.CURR_STS_DESC AS 'Current Status Description',
(CASE WHEN A.WFFSTSDTF = '0001-01-01' THEN NULL ELSE A.WFFSTSDTF END) AS 'Date Last Status Changed by Firm',
(CASE WHEN A.WFFSTSDTF = '0001-01-01' THEN NULL ELSE A.WFFSTSDTF END) AS 'Date Last Status Received on Pipeway', A.LST_NTE_CD AS 'Last Note Code',
A.NTE_DESC AS 'Last Note Code Description',
(CASE WHEN A.LST_NTE_DT = '0001-01-01' THEN NULL ELSE A.LST_NTE_DT END) AS 'Last Note Date', A.LST_NTE_CD AS 'Last Activity Code', A.NTE_DESC AS 'Last Activity Description',
(CASE WHEN A.LST_ACT_DTE = '0001-01-01' THEN NULL ELSE A.LST_ACT_DTE END) AS 'Last Activity Date', A.WFDATEA AS 'Original Plmt Date',
(CASE WHEN A.WFDATEF = '0001-01-01' THEN NULL ELSE A.WFDATEF END) AS 'Date Placed at Firm', A.ORG_PL_AMT AS 'Original Plmt Amount',
A.ORGCODE AS 'Client Code', CONCAT_WS(' ', D.RCLNM1, D.RCLNM2) AS 'Client Name', A.CLT_CDE AS 'Port Code', A.CLNT_DESC AS 'Portfolio Description',
A.CURR_ATTY_CD AS 'Current Firm Code',
A.CURR_ATTY_NME AS 'Current Firm Name', A.DBT_ST_RES AS 'ST', (CASE WHEN A.LST_PYMT_DT_A = '0001-01-01' THEN NULL ELSE A.LST_PYMT_DT_A END) AS 'Last Pay Date',
A.PYMT_TRAN_CD AS 'Last Pay Trans Code', B.SYSDESC AS 'Last Pay Trans Code Desc', A.LST_PYMT_AMT AS 'Last Amount Recovered', A.TTL_PD_ACCT AS 'Total Amount Recovered',
CAST(A.DTE_LST_CC AS DATE) AS 'Last Court Cost Date',
A.TTL_CC AS 'Total Court Cost', (CASE WHEN A.JGMT_DTE = '0001-01-01' THEN NULL ELSE A.JGMT_DTE END) AS 'Judg Date', A.JGMT_AMT AS 'Judg Amt',
A.JGMT_INT_RTE AS 'Judg Int Rate', A.JGMT_CSE_NBR AS 'Case Number', A.CRT_NME AS 'Court', CAST(C.PROCDT AS DATE) 'Date Last Balance Received From Firm',
C.HTOTBAL AS 'Last Balance Received From Firm',
(CASE WHEN  A.CURR_STS_CD IN (93,994) THEN 'SIF'
      WHEN A.CURR_STS_CD = 995 THEN 'PIF' ELSE NULL END) AS 'SIF/PIF',
A.HEMPINFO AS 'POE Flag', E.TRANDTE AS 'Date POE Updated',
(CASE WHEN A.HAJDMT = 1 THEN 'Judgment'
      WHEN A.HSERV = 1 THEN 'Service'
      WHEN A.HSUIT = 1 THEN 'Suit Filed'
      WHEN A.H1DLTR = 1 THEN '1st Demand Letter' ELSE 'None' END) AS 'Current Mile Marker',
A.WFIMGPLFG AS 'Placement Documents', A.WFIMGJDFG AS 'Judgment Copy',
(CASE WHEN A.HAJDMT = 1 THEN WFJDEPA
      WHEN A.HSERV = 1 THEN WFSERVEA
      WHEN A.HSUIT = 1 THEN WFSUITEA
      WHEN A.H1DLTR = 1 THEN WF1DLTREA ELSE 0 END)  AS 'Days to Current Mile Marker',
(CASE WHEN A.HAJDMT = 1 THEN DATEDIFF(CURRENT_DATE(), JGMT_DTE)
      WHEN A.HSERV = 1 THEN DATEDIFF(CURRENT_DATE(), WFSERVDT)
      WHEN A.HSUIT = 1 THEN DATEDIFF(CURRENT_DATE(), WFSUITDT)
      WHEN A.H1DLTR = 1 THEN DATEDIFF(CURRENT_DATE(), WF1DLTRDT) ELSE 0 END) AS 'Days Since Current Mile Marker',
A.H1DLTR AS '1DL Sent Flag',
(CASE WHEN A.WF1DLTRDT <> '0001-01-01' THEN A.WF1DLTRDT ELSE NULL END) AS 'Date 1DL Sent',
A.WF1DLTREA AS 'Days to 1DL Sent',
(CASE WHEN A.WF1DLTRDT <> '0001-01-01' THEN DATEDIFF(CURRENT_DATE(), A.WF1DLTRDT) ELSE NULL END) AS 'Days Since 1DL Sent',
A.HSUIT AS 'Suit Filed Flag',
(CASE WHEN A.WFSUITDT <> '0001-01-01' THEN A.WFSUITDT ELSE NULL END) AS 'Reported Date Suit Filed',
A.WFSUITEA AS 'Days to Suit Filed',
(CASE WHEN A.WFSUITDT <> '0001-01-01' THEN DATEDIFF(CURRENT_DATE(), A.WFSUITDT) ELSE NULL END) AS 'Days Since 1st Filed',
A.HSERV AS 'Service Flag',
(CASE WHEN A.WFSERVDT <> '0001-01-01' THEN A.WFSERVDT ELSE NULL END) AS 'Date of Service',
A.WFSERVEA AS 'Days to Service',
(CASE WHEN A.WFSERVDT <> '0001-01-01' THEN DATEDIFF(CURRENT_DATE(), A.WFSERVDT) ELSE NULL END) AS 'Days Since Service',
A.HAJDMT AS 'Judg Flag',
(CASE WHEN A.JGMT_DTE <> '0001-01-01' THEN A.JGMT_DTE ELSE NULL END) AS 'Date of Judgmt',
A.WFJDEPA AS 'Days to Judgmt',
(CASE WHEN A.JGMT_DTE <> '0001-01-01' THEN DATEDIFF(CURRENT_DATE(), A.JGMT_DTE) ELSE NULL END) AS 'Days Since Judgmt',
A.HACL AS 'Closed Flag',
(CASE WHEN A.WFLCLDT <> '0001-01-01' THEN A.WFLCLDT ELSE NULL END) AS 'Date Closed',
A.WFLCLDYA AS 'Days to Closed',
(CASE WHEN A.WFLCLDT <> '0001-01-01' THEN DATEDIFF(CURRENT_DATE(), A.WFLCLDT) ELSE NULL END) AS 'Days Since Closed',
A.HALIEN AS 'Lien Flag',
(CASE WHEN A.WFLIENDT = '0001-01-01' THEN NULL ELSE A.WFLIENDT END) AS 'Date Lien',
/* AH06052026 - Handle invalid charge off dates to prevent Excel invalid date conversion */
CASE
    WHEN A.RMSDTCHGDT IN ('00000000','0','00010101')
         OR A.RMSDTCHGDT IS NULL
         OR A.RMSDTCHGDT = ''
    THEN NULL
    ELSE STR_TO_DATE(A.RMSDTCHGDT, '%Y%m%d')
END AS RMSDTCHGDT,
A.WFORGPRBAL
FROM HSFLCLNTWF A
LEFT JOIN RMSPSYSCDE B ON A.PYMT_TRAN_CD = B.SYSCODETYP AND B.RMSRECTYPE = 1
LEFT JOIN HUBRELBAGD C ON A.ACCT_NUM = C.RACCTNUM
LEFT JOIN RMRMCLNM_REQ_DATA D ON A.ORGCODE = D.RMSBRGLVL2
LEFT JOIN (
           SELECT FILENUM, MAX(CAST(TRANDTE AS DATE)) AS TRANDTE
           FROM HSTEMPCHG
           WHERE FLDCODE = 'MASENM'
           GROUP BY FILENUM
           ) E ON A.WFFILNUM = E.FILENUM
WHERE ".$where."
AND A.CURR_STS_CD LIKE '9%'";



$result = mysqli_query($conn,$query);


$writer = new XLSXWriter();


$header = array(
			    "Name" => "string",
				"Firm File Number" => "string",
				"Account Number" => "string",
				"Employer Name" => "string",
				"Pipe Status" => "string",
				"Cur Sts Code" => "string",
				"Current Status Description" => "string",
				"Date Last Status Changed by Firm" => "MM-DD-YYYY",
				"Date Last Status Received on Pipeway" => "MM-DD-YYYY",
				"Last Note Code" => "string",
				"Last Note Code Description" => "string",
				"Last Note Date" => "MM-DD-YYYY",
				"Last Activity Code" => "string",
				"Last Activity Description" => "string",
				"Last Activity Date" => "MM-DD-YYYY",
				"Original Plmt Date" => "MM-DD-YYYY",
				"Date Placed at Firm" => "MM-DD-YYYY",
				"Original Plmt Amount" => "dollar",
				"Client Code" => "string",
				"Client Name" => "string",
				"Port Code" => "string",
				"Portfolio Description" => "string",
				"Current Firm Code" => "string",
				"Current Firm Name" => "string",
				"ST" => "string",
				"Last Pay Date" => "MM-DD-YYYY",
				"Last Pay Trans Code" => "string",
				"Last Pay Trans Code Desc" => "string",
				"Last Amount Recovered" => "dollar",
				"Total Amount Recovered" => "dollar",
				"Last Court Cost Date" => "MM-DD-YYYY",
				"Total Court Cost" => "dollar",
				"Judg Date" => "MM-DD-YYYY",
				"Judg Amt" => "dollar",
				"Judg Int Rate" => "dollar",
				"Case Number" => "string",
				"Court" => "string",
				"Date Last Balance Received From Firm" => "MM-DD-YYYY",
				"Last Balance Received From Firm" => "string",
				"SIF/PIF" => "string",
				"POE Flag" => "integer",
				"Date POE Updated" => "MM-DD-YYYY",
				"Current Mile Marker" => "string",
				"Placement Documents" => "string",
				"Judgment Copy" => "string",
				"Days to Current Mile Marker" => "integer",
				"Days Since Current Mile Marker" => "integer",
				"1DL Sent Flag" => "integer",
				"Date 1DL Sent" => "MM-DD-YYYY",
				"Days to 1DL Sent" => "integer",
				"Days Since 1DL Sent" => "integer",
				"Suit Filed Flag" => "integer",
				"Reported Date Suit Filed" => "MM-DD-YYYY",
				"Days to Suit Filed" => "integer",
				"Days Since 1st Filed" => "integer",
				"Service Flag" => "integer",
				"Date of Service" => "MM-DD-YYYY",
				"Days to Service" => "integer",
				"Days Since Service" => "integer",
				"Judg Flag" => "integer",
				"Date of Judgmt" => "MM-DD-YYYY",
				"Days to Judgmt" => "integer",
				"Days Since Judgmt" => "integer",
				"Closed Flag" => "integer",
				"Date Closed" => "MM-DD-YYYY",
				"Days to Closed" => "integer",
				"Days Since Closed" => "integer",
				"Lien Flag" => "integer",
				"Date Lien" => "MM-DD-YYYY",
				"Charge Off Date" => "MM-DD-YYYY",
				"Charge Off Amount" => "dollar",
				);


		      $style = array(
	                'font-style' => 'bold', 'fill'=>'#eee', 'halign'=>'center', 'border'=>'left,right,top,bottom', 'widths'=>[25,25,25,25,25,25,40,40,40,25,40,25,25,36,25,25,25,25,25,25,25,40,25,40,15,25,25,30,25,40,25,25,25,25,25,25,25,40,40,25,25,25,25,25,25,30,30,35,35,25,25,25,25,25,25,25,25,25,25,25,25,25,25,25,25,25,25,25,25,25,25]
	              );
		    
		    // $array = array();
		    $writer->writeSheetHeader('ClosedAccounts', $header, $style);
		   
		    while ($row = mysqli_fetch_assoc($result))
		    {
		         //print_r($row); exit;
		        $writer->writeSheetRow('ClosedAccounts', $row);
		    }

		   //print_r($row); exit();

		    //$writer->writeSheet($array,'Sheet1', $header);//or write the whole sheet in 1 call    

		    // $writer->writeToStdOut();


		    
		    $writer->writeToFile(str_replace(__FILE__,'/var/www/html/bi/dist/Report/uploadsexcel/'.$filename.'',__FILE__));

		   echo json_encode(array('status' => 1));
		    //echo $writer->writeToString();
		    exit(0);

}

else if($accounts == 'open,close' && $_POST['outputFormat'] == '2')
 {
    if($_POST['userType'] == 3)
  {
    $where = "A.ORGCODE = '".$_POST['client']."'";
  }
  else if ($_POST['userType'] == 2 || $_POST['userType'] == 4)
  {
    $where = "A.CURR_ATTY_CD = '".$_POST['client']."' ";
  }


   $filename = "openInventoryReport.xlsx";

    header('Content-disposition: attachment; filename="'.XLSXWriter::sanitize_filename($filename).'"');
    header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    header('Content-Transfer-Encoding: binary');
    header('Cache-Control: must-revalidate');
    header('Pragma: public'); 


  $query = "SELECT A.WFNAME AS 'Name', A.WFFIRMFILE AS 'Firm File Number', A.ACCT_NUM AS 'Account Number', A.MSTPOENAM1 AS 'Employer Name',
(CASE WHEN A.WFPIPEB = 1 THEN 'BLUE'
      WHEN A.WFPIPEO = 1 THEN 'ORANGE'
      WHEN A.WFPIPEG = 1 THEN 'GREEN'
      WHEN A.WFPIPER = 1 THEN 'RED'
      WHEN A.WFPIPEX = 1 THEN 'PURPLE' ELSE 'WHITE' END) AS 'Pipe Status', A.CURR_STS_CD AS 'Cur Sts Code', A.CURR_STS_DESC AS 'Current Status Description',
(CASE WHEN A.WFFSTSDTF = '0001-01-01' THEN NULL ELSE A.WFFSTSDTF END) AS 'Date Last Status Changed by Firm',
(CASE WHEN A.WFFSTSDTF = '0001-01-01' THEN NULL ELSE A.WFFSTSDTF END) AS 'Date Last Status Received on Pipeway', A.LST_NTE_CD AS 'Last Note Code',
A.NTE_DESC AS 'Last Note Code Description',
(CASE WHEN A.LST_NTE_DT = '0001-01-01' THEN NULL ELSE A.LST_NTE_DT END) AS 'Last Note Date', A.LST_NTE_CD AS 'Last Activity Code', A.NTE_DESC AS 'Last Activity Description',
(CASE WHEN A.LST_ACT_DTE = '0001-01-01' THEN NULL ELSE A.LST_ACT_DTE END) AS 'Last Activity Date', A.WFDATEA AS 'Original Plmt Date',
(CASE WHEN A.WFDATEF = '0001-01-01' THEN NULL ELSE A.WFDATEF END) AS 'Date Placed at Firm', A.ORG_PL_AMT AS 'Original Plmt Amount',
A.ORGCODE AS 'Client Code', CONCAT_WS(' ', D.RCLNM1, D.RCLNM2) AS 'Client Name', A.CLT_CDE AS 'Port Code', A.CLNT_DESC AS 'Portfolio Description',
A.CURR_ATTY_CD AS 'Current Firm Code',
A.CURR_ATTY_NME AS 'Current Firm Name', A.DBT_ST_RES AS 'ST', (CASE WHEN A.LST_PYMT_DT_A = '0001-01-01' THEN NULL ELSE A.LST_PYMT_DT_A END) AS 'Last Pay Date',
A.PYMT_TRAN_CD AS 'Last Pay Trans Code', B.SYSDESC AS 'Last Pay Trans Code Desc', A.LST_PYMT_AMT AS 'Last Amount Recovered', A.TTL_PD_ACCT AS 'Total Amount Recovered',
CAST(A.DTE_LST_CC AS DATE) AS 'Last Court Cost Date',
A.TTL_CC AS 'Total Court Cost', (CASE WHEN A.JGMT_DTE = '0001-01-01' THEN NULL ELSE A.JGMT_DTE END) AS 'Judg Date', A.JGMT_AMT AS 'Judg Amt',
A.JGMT_INT_RTE AS 'Judg Int Rate', A.JGMT_CSE_NBR AS 'Case Number', A.CRT_NME AS 'Court', CAST(C.PROCDT AS DATE) 'Date Last Balance Received From Firm',
C.HTOTBAL AS 'Last Balance Received From Firm',
(CASE WHEN  A.CURR_STS_CD IN (93,994) THEN 'SIF'
      WHEN A.CURR_STS_CD = 995 THEN 'PIF' ELSE NULL END) AS 'SIF/PIF',
A.HEMPINFO AS 'POE Flag', E.TRANDTE AS 'Date POE Updated',
(CASE WHEN A.HAJDMT = 1 THEN 'Judgment'
      WHEN A.HSERV = 1 THEN 'Service'
      WHEN A.HSUIT = 1 THEN 'Suit Filed'
      WHEN A.H1DLTR = 1 THEN '1st Demand Letter' ELSE 'None' END) AS 'Current Mile Marker',
A.WFIMGPLFG AS 'Placement Documents', A.WFIMGJDFG AS 'Judgment Copy',
(CASE WHEN A.HAJDMT = 1 THEN WFJDEPA
      WHEN A.HSERV = 1 THEN WFSERVEA
      WHEN A.HSUIT = 1 THEN WFSUITEA
      WHEN A.H1DLTR = 1 THEN WF1DLTREA ELSE 0 END)  AS 'Days to Current Mile Marker',
(CASE WHEN A.HAJDMT = 1 THEN DATEDIFF(CURRENT_DATE(), JGMT_DTE)
      WHEN A.HSERV = 1 THEN DATEDIFF(CURRENT_DATE(), WFSERVDT)
      WHEN A.HSUIT = 1 THEN DATEDIFF(CURRENT_DATE(), WFSUITDT)
      WHEN A.H1DLTR = 1 THEN DATEDIFF(CURRENT_DATE(), WF1DLTRDT) ELSE 0 END) AS 'Days Since Current Mile Marker',
A.H1DLTR AS '1DL Sent Flag',
(CASE WHEN A.WF1DLTRDT <> '0001-01-01' THEN A.WF1DLTRDT ELSE NULL END) AS 'Date 1DL Sent',
A.WF1DLTREA AS 'Days to 1DL Sent',
(CASE WHEN A.WF1DLTRDT <> '0001-01-01' THEN DATEDIFF(CURRENT_DATE(), A.WF1DLTRDT) ELSE NULL END) AS 'Days Since 1DL Sent',
A.HSUIT AS 'Suit Filed Flag',
(CASE WHEN A.WFSUITDT <> '0001-01-01' THEN A.WFSUITDT ELSE NULL END) AS 'Reported Date Suit Filed',
A.WFSUITEA AS 'Days to Suit Filed',
(CASE WHEN A.WFSUITDT <> '0001-01-01' THEN DATEDIFF(CURRENT_DATE(), A.WFSUITDT) ELSE NULL END) AS 'Days Since 1st Filed',
A.HSERV AS 'Service Flag',
(CASE WHEN A.WFSERVDT <> '0001-01-01' THEN A.WFSERVDT ELSE NULL END) AS 'Date of Service',
A.WFSERVEA AS 'Days to Service',
(CASE WHEN A.WFSERVDT <> '0001-01-01' THEN DATEDIFF(CURRENT_DATE(), A.WFSERVDT) ELSE NULL END) AS 'Days Since Service',
A.HAJDMT AS 'Judg Flag',
(CASE WHEN A.JGMT_DTE <> '0001-01-01' THEN A.JGMT_DTE ELSE NULL END) AS 'Date of Judgmt',
A.WFJDEPA AS 'Days to Judgmt',
(CASE WHEN A.JGMT_DTE <> '0001-01-01' THEN DATEDIFF(CURRENT_DATE(), A.JGMT_DTE) ELSE NULL END) AS 'Days Since Judgmt',
A.HACL AS 'Closed Flag',
(CASE WHEN A.WFLCLDT <> '0001-01-01' THEN A.WFLCLDT ELSE NULL END) AS 'Date Closed',
A.WFLCLDYA AS 'Days to Closed',
(CASE WHEN A.WFLCLDT <> '0001-01-01' THEN DATEDIFF(CURRENT_DATE(), A.WFLCLDT) ELSE NULL END) AS 'Days Since Closed',
A.HALIEN AS 'Lien Flag',
(CASE WHEN A.WFLIENDT = '0001-01-01' THEN NULL ELSE A.WFLIENDT END) AS 'Date Lien',
/* AH06052026 - Handle invalid charge off dates to prevent Excel invalid date conversion */
CASE
    WHEN A.RMSDTCHGDT IN ('00000000','0','00010101')
         OR A.RMSDTCHGDT IS NULL
         OR A.RMSDTCHGDT = ''
    THEN NULL
    ELSE STR_TO_DATE(A.RMSDTCHGDT, '%Y%m%d')
END AS RMSDTCHGDT,
A.WFORGPRBAL 
FROM HSFLCLNTWF A
LEFT JOIN RMSPSYSCDE B ON A.PYMT_TRAN_CD = B.SYSCODETYP AND B.RMSRECTYPE = 1
LEFT JOIN HUBRELBAGD C ON A.ACCT_NUM = C.RACCTNUM
LEFT JOIN RMRMCLNM_REQ_DATA D ON A.ORGCODE = D.RMSBRGLVL2
LEFT JOIN (
           SELECT FILENUM, MAX(CAST(TRANDTE AS DATE)) AS TRANDTE
           FROM HSTEMPCHG
           WHERE FLDCODE = 'MASENM'
           GROUP BY FILENUM
           ) E ON A.WFFILNUM = E.FILENUM
WHERE ".$where." ";


$result = mysqli_query($conn,$query);

$writer = new XLSXWriter();

$header = array(
			    "Name" => "string",
				"Firm File Number" => "string",
				"Account Number" => "string",
				"Employer Name" => "string",
				"Pipe Status" => "string",
				"Cur Sts Code" => "string",
				"Current Status Description" => "string",
				"Date Last Status Changed by Firm" => "MM-DD-YYYY",
				"Date Last Status Received on Pipeway" => "MM-DD-YYYY",
				"Last Note Code" => "string",
				"Last Note Code Description" => "string",
				"Last Note Date" => "MM-DD-YYYY",
				"Last Activity Code" => "string",
				"Last Activity Description" => "string",
				"Last Activity Date" => "MM-DD-YYYY",
				"Original Plmt Date" => "MM-DD-YYYY",
				"Date Placed at Firm" => "MM-DD-YYYY",
				"Original Plmt Amount" => "dollar",
				"Client Code" => "string",
				"Client Name" => "string",
				"Port Code" => "string",
				"Portfolio Description" => "string",
				"Current Firm Code" => "string",
				"Current Firm Name" => "string",
				"ST" => "string",
				"Last Pay Date" => "MM-DD-YYYY",
				"Last Pay Trans Code" => "string",
				"Last Pay Trans Code Desc" => "string",
				"Last Amount Recovered" => "dollar",
				"Total Amount Recovered" => "dollar",
				"Last Court Cost Date" => "MM-DD-YYYY",
				"Total Court Cost" => "dollar",
				"Judg Date" => "MM-DD-YYYY",
				"Judg Amt" => "dollar",
				"Judg Int Rate" => "dollar",
				"Case Number" => "string",
				"Court" => "string",
				"Date Last Balance Received From Firm" => "MM-DD-YYYY",
				"Last Balance Received From Firm" => "string",
				"SIF/PIF" => "string",
				"POE Flag" => "integer",
				"Date POE Updated" => "MM-DD-YYYY",
				"Current Mile Marker" => "string",
				"Placement Documents" => "string",
				"Judgment Copy" => "string",
				"Days to Current Mile Marker" => "integer",
				"Days Since Current Mile Marker" => "integer",
				"1DL Sent Flag" => "integer",
				"Date 1DL Sent" => "MM-DD-YYYY",
				"Days to 1DL Sent" => "integer",
				"Days Since 1DL Sent" => "integer",
				"Suit Filed Flag" => "integer",
				"Reported Date Suit Filed" => "MM-DD-YYYY",
				"Days to Suit Filed" => "integer",
				"Days Since 1st Filed" => "integer",
				"Service Flag" => "integer",
				"Date of Service" => "MM-DD-YYYY",
				"Days to Service" => "integer",
				"Days Since Service" => "integer",
				"Judg Flag" => "integer",
				"Date of Judgmt" => "MM-DD-YYYY",
				"Days to Judgmt" => "integer",
				"Days Since Judgmt" => "integer",
				"Closed Flag" => "integer",
				"Date Closed" => "MM-DD-YYYY",
				"Days to Closed" => "integer",
				"Days Since Closed" => "integer",
				"Lien Flag" => "integer",
				"Date Lien" => "MM-DD-YYYY",
				"Charge Off Date" => "MM-DD-YYYY",
				"Charge Off Amount" => "dollar",
				);


		      $style = array(
	                'font-style' => 'bold', 'fill'=>'#eee', 'halign'=>'center', 'border'=>'left,right,top,bottom', 'widths'=>[25,25,25,25,25,25,40,40,40,25,40,25,25,36,25,25,25,25,25,25,25,40,25,40,15,25,25,30,25,40,25,25,25,25,25,25,25,40,40,25,25,25,25,25,25,30,30,35,35,25,25,25,25,25,25,25,25,25,25,25,25,25,25,25,25,25,25,25,25,25,25]
	              );
		    
		   
		    $writer->writeSheetHeader('OpenAndClosedAccounts', $header, $style);
		   
		    while ($row = mysqli_fetch_assoc($result))
		    {
		        
		        $writer->writeSheetRow('OpenAndClosedAccounts', $row);
		    }

		  
		    
		    $writer->writeToFile(str_replace(__FILE__,'/var/www/html/bi/dist/Report/uploadsexcel/'.$filename.'',__FILE__));

		   echo json_encode(array('status' => 1));
		   
		    exit(0);
 }
 else if($accounts == 'pending' && $_POST['outputFormat'] == '2')
 {
   if($_POST['userType'] == 3)
  {
    $where = "B.RMSBRGLVL2 = '".$_POST['client']."'";
  }
  else if ($_POST['userType'] == 2 || $_POST['userType'] == 4)
  {
    $where = "A.ATTRNYCODE = '".$_POST['client']."' ";
  }


   $filename = "openInventoryReport.xlsx";

    header('Content-disposition: attachment; filename="'.XLSXWriter::sanitize_filename($filename).'"');
    header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    header('Content-Transfer-Encoding: binary');
    header('Cache-Control: must-revalidate');
    header('Pragma: public'); 


  $query = "SELECT A.RMSACCTNUM AS 'Account Number' , CONCAT_WS(' ', A.RMSCORPNM1, A.RMSCORPNM2) AS 'Name', A.RMSSTATECD AS 'ST', A.RMSACCNTST AS 'Cur Sts Code',
A.ATTRNYCODE AS 'Current Firm Code',  B.RMSBRGLVL2 AS 'Client Code', A.RMSOFFCRCD AS 'Port Code', CAST(A.RMSDATERCV AS DATE) AS 'Original Plmt Date',
CAST(A.RMSDATEASG AS DATE) AS 'Date Placed at Firm', A.LSTCOMMNT1 AS 'Last Comment1', A.LSTCOMMNT2 AS 'Last Comment2', A.LSTCOMMNT3 AS 'Last Comment3',
CAST(A.LSTCMMTDTE AS DATE) AS 'Last Comment Dt', A.MSTPOENAM1 AS 'Employer Name', CAST(A.LSTSTATCHG AS DATE) AS 'Last Status Dt', A.LSTSSTATUS AS 'Last Status Code',
/* AH06052026 - Handle invalid charge off dates to prevent Excel invalid date conversion */
CASE
    WHEN A.RMSDTCHGDT IN ('00000000','0','00010101')
         OR A.RMSDTCHGDT IS NULL
         OR A.RMSDTCHGDT = ''
    THEN NULL
    ELSE STR_TO_DATE(A.RMSDTCHGDT, '%Y%m%d')
END AS RMSDTCHGDT,
A.RMSCHGAMT 
FROM RMSPMASTRP A
INNER JOIN RMSPSYSASN_REQ_DATA B ON A.RMSOFFCRCD = B.RMSOFFCRCD AND B.RMSRECTYPE = '1'
WHERE ".$where." ";


$result = mysqli_query($conn,$query);

$writer = new XLSXWriter();

$header = array(
			    "Account Number" => "string",
				"Name" => "string",
				"ST" => "string",
				"Cur Sts Code" => "string",
				"Current Firm Code" => "string",
				"Client Code" => "string",
				"Port Code" => "string",
				"Original Plmt Date" => "MM-DD-YYYY",
				"Date Placed at Firm" => "MM-DD-YYYY",
				"Last Comment1" => "string",
				"Last Comment2" => "string",
				"Last Comment3" => "string",
				"Last Comment Dt" => "MM-DD-YYYY",
				"Employer Name" => "string",
				"Last Status Dt" => "MM-DD-YYYY",
				"Last Status Code" => "string",
				"Charge Off Date" => "MM-DD-YYYY",
				"Charge Off Amount" => "dollar",
				);


		      $style = array(
		        'font-style' => 'bold', 'fill'=>'#eee', 'halign'=>'center', 'border'=>'left,right,top,bottom', 'widths'=>[25,25,25,25,30,25,25,30,30,25,25,25,25,25,25,25,25,25]
		    );
		    
		   
		    $writer->writeSheetHeader('PendingAccounts', $header, $style);
		   
		    while ($row = mysqli_fetch_assoc($result))
		    {
		        
		        $writer->writeSheetRow('PendingAccounts', $row);
		    }

		  
		    
		    $writer->writeToFile(str_replace(__FILE__,'/var/www/html/bi/dist/Report/uploadsexcel/'.$filename.'',__FILE__));

		   echo json_encode(array('status' => 1));
		   
		    exit(0);
 }

 else if ($accounts == 'open,close,pending' && $_POST['outputFormat'] == '2')
 {

 	if($_POST['userType'] == 3)
  {
  	$where = "A.ORGCODE = '".$_POST['client']."'";
    $where1 = "B.RMSBRGLVL2 = '".$_POST['client']."'";
  }
  else if ($_POST['userType'] == 2 || $_POST['userType'] == 4)
  {
  	$where = "A.CURR_ATTY_CD = '".$_POST['client']."' ";
    $where1 = "A.ATTRNYCODE = '".$_POST['client']."' ";
  }




   $filename = "openInventoryReport.xlsx";

    header('Content-disposition: attachment; filename="'.XLSXWriter::sanitize_filename($filename).'"');
    header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    header('Content-Transfer-Encoding: binary');
    header('Cache-Control: must-revalidate');
    header('Pragma: public'); 


  $query1 = "SELECT A.WFNAME AS 'Name', A.WFFIRMFILE AS 'Firm File Number', A.ACCT_NUM AS 'Account Number', A.MSTPOENAM1 AS 'Employer Name',
(CASE WHEN A.WFPIPEB = 1 THEN 'BLUE'
      WHEN A.WFPIPEO = 1 THEN 'ORANGE'
      WHEN A.WFPIPEG = 1 THEN 'GREEN'
      WHEN A.WFPIPER = 1 THEN 'RED'
      WHEN A.WFPIPEX = 1 THEN 'PURPLE' ELSE 'WHITE' END) AS 'Pipe Status', A.CURR_STS_CD AS 'Cur Sts Code', A.CURR_STS_DESC AS 'Current Status Description',
(CASE WHEN A.WFFSTSDTF = '0001-01-01' THEN NULL ELSE A.WFFSTSDTF END) AS 'Date Last Status Changed by Firm',
(CASE WHEN A.WFFSTSDTF = '0001-01-01' THEN NULL ELSE A.WFFSTSDTF END) AS 'Date Last Status Received on Pipeway', A.LST_NTE_CD AS 'Last Note Code',
A.NTE_DESC AS 'Last Note Code Description',
(CASE WHEN A.LST_NTE_DT = '0001-01-01' THEN NULL ELSE A.LST_NTE_DT END) AS 'Last Note Date', A.LST_NTE_CD AS 'Last Activity Code', A.NTE_DESC AS 'Last Activity Description',
(CASE WHEN A.LST_ACT_DTE = '0001-01-01' THEN NULL ELSE A.LST_ACT_DTE END) AS 'Last Activity Date', A.WFDATEA AS 'Original Plmt Date',
(CASE WHEN A.WFDATEF = '0001-01-01' THEN NULL ELSE A.WFDATEF END) AS 'Date Placed at Firm', A.ORG_PL_AMT AS 'Original Plmt Amount',
A.ORGCODE AS 'Client Code', CONCAT_WS(' ', D.RCLNM1, D.RCLNM2) AS 'Client Name', A.CLT_CDE AS 'Port Code', A.CLNT_DESC AS 'Portfolio Description',
A.CURR_ATTY_CD AS 'Current Firm Code',
A.CURR_ATTY_NME AS 'Current Firm Name', A.DBT_ST_RES AS 'ST', (CASE WHEN A.LST_PYMT_DT_A = '0001-01-01' THEN NULL ELSE A.LST_PYMT_DT_A END) AS 'Last Pay Date',
A.PYMT_TRAN_CD AS 'Last Pay Trans Code', B.SYSDESC AS 'Last Pay Trans Code Desc', A.LST_PYMT_AMT AS 'Last Amount Recovered', A.TTL_PD_ACCT AS 'Total Amount Recovered',
CAST(A.DTE_LST_CC AS DATE) AS 'Last Court Cost Date',
A.TTL_CC AS 'Total Court Cost', (CASE WHEN A.JGMT_DTE = '0001-01-01' THEN NULL ELSE A.JGMT_DTE END) AS 'Judg Date', A.JGMT_AMT AS 'Judg Amt',
A.JGMT_INT_RTE AS 'Judg Int Rate', A.JGMT_CSE_NBR AS 'Case Number', A.CRT_NME AS 'Court', CAST(C.PROCDT AS DATE) 'Date Last Balance Received From Firm',
C.HTOTBAL AS 'Last Balance Received From Firm',
(CASE WHEN  A.CURR_STS_CD IN (93,994) THEN 'SIF'
      WHEN A.CURR_STS_CD = 995 THEN 'PIF' ELSE NULL END) AS 'SIF/PIF',
A.HEMPINFO AS 'POE Flag', E.TRANDTE AS 'Date POE Updated',
(CASE WHEN A.HAJDMT = 1 THEN 'Judgment'
      WHEN A.HSERV = 1 THEN 'Service'
      WHEN A.HSUIT = 1 THEN 'Suit Filed'
      WHEN A.H1DLTR = 1 THEN '1st Demand Letter' ELSE 'None' END) AS 'Current Mile Marker',
A.WFIMGPLFG AS 'Placement Documents', A.WFIMGJDFG AS 'Judgment Copy',
(CASE WHEN A.HAJDMT = 1 THEN WFJDEPA
      WHEN A.HSERV = 1 THEN WFSERVEA
      WHEN A.HSUIT = 1 THEN WFSUITEA
      WHEN A.H1DLTR = 1 THEN WF1DLTREA ELSE 0 END)  AS 'Days to Current Mile Marker',
(CASE WHEN A.HAJDMT = 1 THEN DATEDIFF(CURRENT_DATE(), JGMT_DTE)
      WHEN A.HSERV = 1 THEN DATEDIFF(CURRENT_DATE(), WFSERVDT)
      WHEN A.HSUIT = 1 THEN DATEDIFF(CURRENT_DATE(), WFSUITDT)
      WHEN A.H1DLTR = 1 THEN DATEDIFF(CURRENT_DATE(), WF1DLTRDT) ELSE 0 END) AS 'Days Since Current Mile Marker',
A.H1DLTR AS '1DL Sent Flag',
(CASE WHEN A.WF1DLTRDT <> '0001-01-01' THEN A.WF1DLTRDT ELSE NULL END) AS 'Date 1DL Sent',
A.WF1DLTREA AS 'Days to 1DL Sent',
(CASE WHEN A.WF1DLTRDT <> '0001-01-01' THEN DATEDIFF(CURRENT_DATE(), A.WF1DLTRDT) ELSE NULL END) AS 'Days Since 1DL Sent',
A.HSUIT AS 'Suit Filed Flag',
(CASE WHEN A.WFSUITDT <> '0001-01-01' THEN A.WFSUITDT ELSE NULL END) AS 'Reported Date Suit Filed',
A.WFSUITEA AS 'Days to Suit Filed',
(CASE WHEN A.WFSUITDT <> '0001-01-01' THEN DATEDIFF(CURRENT_DATE(), A.WFSUITDT) ELSE NULL END) AS 'Days Since 1st Filed',
A.HSERV AS 'Service Flag',
(CASE WHEN A.WFSERVDT <> '0001-01-01' THEN A.WFSERVDT ELSE NULL END) AS 'Date of Service',
A.WFSERVEA AS 'Days to Service',
(CASE WHEN A.WFSERVDT <> '0001-01-01' THEN DATEDIFF(CURRENT_DATE(), A.WFSERVDT) ELSE NULL END) AS 'Days Since Service',
A.HAJDMT AS 'Judg Flag',
(CASE WHEN A.JGMT_DTE <> '0001-01-01' THEN A.JGMT_DTE ELSE NULL END) AS 'Date of Judgmt',
A.WFJDEPA AS 'Days to Judgmt',
(CASE WHEN A.JGMT_DTE <> '0001-01-01' THEN DATEDIFF(CURRENT_DATE(), A.JGMT_DTE) ELSE NULL END) AS 'Days Since Judgmt',
A.HACL AS 'Closed Flag',
(CASE WHEN A.WFLCLDT <> '0001-01-01' THEN A.WFLCLDT ELSE NULL END) AS 'Date Closed',
A.WFLCLDYA AS 'Days to Closed',
(CASE WHEN A.WFLCLDT <> '0001-01-01' THEN DATEDIFF(CURRENT_DATE(), A.WFLCLDT) ELSE NULL END) AS 'Days Since Closed',
A.HALIEN AS 'Lien Flag',
(CASE WHEN A.WFLIENDT = '0001-01-01' THEN NULL ELSE A.WFLIENDT END) AS 'Date Lien',
/* AH06052026 - Handle invalid charge off dates to prevent Excel invalid date conversion */
CASE
    WHEN A.RMSDTCHGDT IN ('00000000','0','00010101')
         OR A.RMSDTCHGDT IS NULL
         OR A.RMSDTCHGDT = ''
    THEN NULL
    ELSE STR_TO_DATE(A.RMSDTCHGDT, '%Y%m%d')
END AS RMSDTCHGDT,
A.WFORGPRBAL 
FROM HSFLCLNTWF A
LEFT JOIN RMSPSYSCDE B ON A.PYMT_TRAN_CD = B.SYSCODETYP AND B.RMSRECTYPE = 1
LEFT JOIN HUBRELBAGD C ON A.ACCT_NUM = C.RACCTNUM
LEFT JOIN RMRMCLNM_REQ_DATA D ON A.ORGCODE = D.RMSBRGLVL2
LEFT JOIN (
           SELECT FILENUM, MAX(CAST(TRANDTE AS DATE)) AS TRANDTE
           FROM HSTEMPCHG
           WHERE FLDCODE = 'MASENM'
           GROUP BY FILENUM
           ) E ON A.WFFILNUM = E.FILENUM
WHERE ".$where." ";



$query2 = "SELECT A.RMSACCTNUM AS 'Account Number' , CONCAT_WS(' ', A.RMSCORPNM1, A.RMSCORPNM2) AS 'Name', A.RMSSTATECD AS 'ST', A.RMSACCNTST AS 'Cur Sts Code',
A.ATTRNYCODE AS 'Current Firm Code',  B.RMSBRGLVL2 AS 'Client Code', A.RMSOFFCRCD AS 'Port Code', CAST(A.RMSDATERCV AS DATE) AS 'Original Plmt Date',
CAST(A.RMSDATEASG AS DATE) AS 'Date Placed at Firm', A.LSTCOMMNT1 AS 'Last Comment1', A.LSTCOMMNT2 AS 'Last Comment2', A.LSTCOMMNT3 AS 'Last Comment3',
CAST(A.LSTCMMTDTE AS DATE) AS 'Last Comment Dt', A.MSTPOENAM1 AS 'Employer Name', CAST(A.LSTSTATCHG AS DATE) AS 'Last Status Dt', A.LSTSSTATUS AS 'Last Status Code',
/* AH06052026 - Handle invalid charge off dates to prevent Excel invalid date conversion */
CASE
    WHEN A.RMSDTCHGDT IN ('00000000','0','00010101')
         OR A.RMSDTCHGDT IS NULL
         OR A.RMSDTCHGDT = ''
    THEN NULL
    ELSE STR_TO_DATE(A.RMSDTCHGDT, '%Y%m%d')
END AS RMSDTCHGDT,
A.RMSCHGAMT 
FROM RMSPMASTRP A
INNER JOIN RMSPSYSASN_REQ_DATA B ON A.RMSOFFCRCD = B.RMSOFFCRCD AND B.RMSRECTYPE = '1'
WHERE ".$where1." ";


$result1 = mysqli_query($conn,$query1);
$result2 = mysqli_query($conn,$query2);

$writer = new XLSXWriter();

$header1 = array(
			    "Name" => "string",
				"Firm File Number" => "string",
				"Account Number" => "string",
				"Employer Name" => "string",
				"Pipe Status" => "string",
				"Cur Sts Code" => "string",
				"Current Status Description" => "string",
				"Date Last Status Changed by Firm" => "MM-DD-YYYY",
				"Date Last Status Received on Pipeway" => "MM-DD-YYYY",
				"Last Note Code" => "string",
				"Last Note Code Description" => "string",
				"Last Note Date" => "MM-DD-YYYY",
				"Last Activity Code" => "string",
				"Last Activity Description" => "string",
				"Last Activity Date" => "MM-DD-YYYY",
				"Original Plmt Date" => "MM-DD-YYYY",
				"Date Placed at Firm" => "MM-DD-YYYY",
				"Original Plmt Amount" => "dollar",
				"Client Code" => "string",
				"Client Name" => "string",
				"Port Code" => "string",
				"Portfolio Description" => "string",
				"Current Firm Code" => "string",
				"Current Firm Name" => "string",
				"ST" => "string",
				"Last Pay Date" => "MM-DD-YYYY",
				"Last Pay Trans Code" => "string",
				"Last Pay Trans Code Desc" => "string",
				"Last Amount Recovered" => "dollar",
				"Total Amount Recovered" => "dollar",
				"Last Court Cost Date" => "MM-DD-YYYY",
				"Total Court Cost" => "dollar",
				"Judg Date" => "MM-DD-YYYY",
				"Judg Amt" => "dollar",
				"Judg Int Rate" => "dollar",
				"Case Number" => "string",
				"Court" => "string",
				"Date Last Balance Received From Firm" => "MM-DD-YYYY",
				"Last Balance Received From Firm" => "string",
				"SIF/PIF" => "string",
				"POE Flag" => "integer",
				"Date POE Updated" => "MM-DD-YYYY",
				"Current Mile Marker" => "string",
				"Placement Documents" => "string",
				"Judgment Copy" => "string",
				"Days to Current Mile Marker" => "integer",
				"Days Since Current Mile Marker" => "integer",
				"1DL Sent Flag" => "integer",
				"Date 1DL Sent" => "MM-DD-YYYY",
				"Days to 1DL Sent" => "integer",
				"Days Since 1DL Sent" => "integer",
				"Suit Filed Flag" => "integer",
				"Reported Date Suit Filed" => "MM-DD-YYYY",
				"Days to Suit Filed" => "integer",
				"Days Since 1st Filed" => "integer",
				"Service Flag" => "integer",
				"Date of Service" => "MM-DD-YYYY",
				"Days to Service" => "integer",
				"Days Since Service" => "integer",
				"Judg Flag" => "integer",
				"Date of Judgmt" => "MM-DD-YYYY",
				"Days to Judgmt" => "integer",
				"Days Since Judgmt" => "integer",
				"Closed Flag" => "integer",
				"Date Closed" => "MM-DD-YYYY",
				"Days to Closed" => "integer",
				"Days Since Closed" => "integer",
				"Lien Flag" => "integer",
				"Date Lien" => "MM-DD-YYYY",
				"Charge Off Date" => "MM-DD-YYYY",
				"Charge Off Amount" => "dollar",
				);


$header2 = array(
			    "Account Number" => "string",
				"Name" => "string",
				"ST" => "string",
				"Cur Sts Code" => "string",
				"Current Firm Code" => "string",
				"Client Code" => "string",
				"Port Code" => "string",
				"Original Plmt Date" => "MM-DD-YYYY",
				"Date Placed at Firm" => "MM-DD-YYYY",
				"Last Comment1" => "string",
				"Last Comment2" => "string",
				"Last Comment3" => "string",
				"Last Comment Dt" => "MM-DD-YYYY",
				"Employer Name" => "string",
				"Last Status Dt" => "MM-DD-YYYY",
				"Last Status Code" => "string",
				"Charge Off Date" => "MM-DD-YYYY",
				"Charge Off Amount" => "dollar",
				);


              
              $style = array(
	                'font-style' => 'bold', 'fill'=>'#eee', 'halign'=>'center', 'border'=>'left,right,top,bottom', 'widths'=>[25,25,25,25,25,25,40,40,40,25,40,25,25,36,25,25,25,25,25,25,25,40,25,40,15,25,25,30,25,40,25,25,25,25,25,25,25,40,40,25,25,25,25,25,25,30,30,35,35,25,25,25,25,25,25,25,25,25,25,25,25,25,25,25,25,25,25,25,25,25,25]
	              );

		      $style1 = array(
		        'font-style' => 'bold', 'fill'=>'#eee', 'halign'=>'center', 'border'=>'left,right,top,bottom', 'widths'=>[25,25,25,25,30,25,25,30,30,25,25,25,25,25,25,25,25,25]
		    );
		    
		   
		     $writer->writeSheetHeader('OpenAndClosedAccounts', $header1, $style);
             $writer->writeSheetHeader('PendingAccounts', $header2, $style1);
		   
		    while ($row = mysqli_fetch_assoc($result1))
		    {
		        
		        $writer->writeSheetRow('OpenAndClosedAccounts', $row);
		    }

		    while ($row2 = mysqli_fetch_assoc($result2))
		    {
		        
		        $writer->writeSheetRow('PendingAccounts', $row2);
		    }

		  
		    
		    $writer->writeToFile(str_replace(__FILE__,'/var/www/html/bi/dist/Report/uploadsexcel/'.$filename.'',__FILE__));

		   echo json_encode(array('status' => 1));
		   
		    exit(0);

 }

}

?>
