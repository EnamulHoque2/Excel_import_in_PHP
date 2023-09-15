<?php
  $dbhost = 'localhost';
  $dbuser = 'dbuser name';
  $dbpass = 'password';
  $dbname = 'db name';
  $con = mysqli_connect($dbhost, $dbuser, $dbpass,$dbname);
 
 if(isset($_POST['submit'])){
	$file=$_FILES['doc']['tmp_name'];
	
	$ext=pathinfo($_FILES['doc']['name'],PATHINFO_EXTENSION);
	if($ext=='xlsx'){
		require('PHPExcel/PHPExcel.php');
		require('PHPExcel/PHPExcel/IOFactory.php');
		
		
		$obj=PHPExcel_IOFactory::load($file);
		foreach($obj->getWorksheetIterator() as $sheet){
			$getHighestRow=$sheet->getHighestRow();
			for($i=2;$i<=$getHighestRow;$i++){
				$col1=$sheet->getCellByColumnAndRow(0,$i)->getValue();
				$col2=$sheet->getCellByColumnAndRow(1,$i)->getValue();
				$col3=$sheet->getCellByColumnAndRow(2,$i)->getValue();
				$col4=$sheet->getCellByColumnAndRow(3,$i)->getValue();
				$col5=$sheet->getCellByColumnAndRow(4,$i)->getValue();
				$col6=$sheet->getCellByColumnAndRow(5,$i)->getValue();
				$col7=$sheet->getCellByColumnAndRow(6,$i)->getValue();
				$col8=$sheet->getCellByColumnAndRow(7,$i)->getValue();
				$col9=$sheet->getCellByColumnAndRow(8,$i)->getValue();
				$col10=$sheet->getCellByColumnAndRow(9,$i)->getValue();
				$col11=$sheet->getCellByColumnAndRow(10,$i)->getValue();
				$col12=$sheet->getCellByColumnAndRow(11,$i)->getValue();
				if(PHPExcel_Shared_Date::isDateTime($sheet->getCellByColumnAndRow(12,$i))) {
				    $pdate1Value = $sheet->getCellByColumnAndRow(12,$i)->getValue();
				    $pdate1timestamp = PHPExcel_Shared_Date::ExcelToPHP($pdate1Value);
				    $col13= date('Y-m-d', $pdate1timestamp);
				}
				$col14=$sheet->getCellByColumnAndRow(13,$i)->getValue();
				$col15=$sheet->getCellByColumnAndRow(14,$i)->getValue();
				$col16=$sheet->getCellByColumnAndRow(15,$i)->getValue();
				$col17=$sheet->getCellByColumnAndRow(16,$i)->getValue();
				$col18=$sheet->getCellByColumnAndRow(17,$i)->getValue();
				$col19=$sheet->getCellByColumnAndRow(18,$i)->getValue();
				if(PHPExcel_Shared_Date::isDateTime($sheet->getCellByColumnAndRow(19,$i))) {
				    $dobValue = $sheet->getCellByColumnAndRow(19,$i)->getValue();
				    $dobimestamp = PHPExcel_Shared_Date::ExcelToPHP($dobValue);
				    $col20= date('Y-m-d', $dobimestamp);
				}
				$col30=$sheet->getCellByColumnAndRow(20,$i)->getValue();
				

				if($stu_name!=''){
					mysqli_query($con,"insert into tablename(col1,col2,col3,col4,col5,col6,session,course,email,class,amount,trn,pdate,pstatus,pmode,gatetrn,paction,dept_roll,gen,dob,cat) values('$stu_name','$father','$mother','$reg_no','$mobile','$product','$session','$course','$email','$class','$amount','$trn','$pdate','$pstatus','$pmode','$gatetrn','$paction','$dept_roll','$gender','$dob','$category')");
				}
			}
		}
	}else{
		echo "Invalid file format";
	}
} 
?>
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Excel data import into database</title>
</head>
<body>
<div class="container">
   <form method="post" enctype="multipart/form-data" style="display:flex;">
         <input type="file" style="margin-bottom:2px;" name="doc"/>
         <button type="submit" class="btn btn-info mx-2" name="submit">Import Excel</button>
    </form>
</div>
</body>
</html>
