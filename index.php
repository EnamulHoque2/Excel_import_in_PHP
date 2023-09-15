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
				$stu_name=$sheet->getCellByColumnAndRow(0,$i)->getValue();
				$father=$sheet->getCellByColumnAndRow(1,$i)->getValue();
				$mother=$sheet->getCellByColumnAndRow(2,$i)->getValue();
				$reg_no=$sheet->getCellByColumnAndRow(3,$i)->getValue();
				$mobile=$sheet->getCellByColumnAndRow(4,$i)->getValue();
				$product=$sheet->getCellByColumnAndRow(5,$i)->getValue();
				$session=$sheet->getCellByColumnAndRow(6,$i)->getValue();
				$course=$sheet->getCellByColumnAndRow(7,$i)->getValue();
				$email=$sheet->getCellByColumnAndRow(8,$i)->getValue();
				$class=$sheet->getCellByColumnAndRow(9,$i)->getValue();
				$amount=$sheet->getCellByColumnAndRow(10,$i)->getValue();
				$trn=$sheet->getCellByColumnAndRow(11,$i)->getValue();
				if(PHPExcel_Shared_Date::isDateTime($sheet->getCellByColumnAndRow(12,$i))) {
				    $pdate1Value = $sheet->getCellByColumnAndRow(12,$i)->getValue();
				    $pdate1timestamp = PHPExcel_Shared_Date::ExcelToPHP($pdate1Value);
				    $pdate= date('Y-m-d', $pdate1timestamp);
				}
				$pstatus=$sheet->getCellByColumnAndRow(13,$i)->getValue();
				$pmode=$sheet->getCellByColumnAndRow(14,$i)->getValue();
				$gatetrn=$sheet->getCellByColumnAndRow(15,$i)->getValue();
				$paction=$sheet->getCellByColumnAndRow(16,$i)->getValue();
				$dept_roll=$sheet->getCellByColumnAndRow(17,$i)->getValue();
				$gender=$sheet->getCellByColumnAndRow(18,$i)->getValue();
				if(PHPExcel_Shared_Date::isDateTime($sheet->getCellByColumnAndRow(19,$i))) {
				    $dobValue = $sheet->getCellByColumnAndRow(19,$i)->getValue();
				    $dobimestamp = PHPExcel_Shared_Date::ExcelToPHP($dobValue);
				    $dob= date('Y-m-d', $dobimestamp);
				}
				$category=$sheet->getCellByColumnAndRow(20,$i)->getValue();
				

				if($stu_name!=''){
					mysqli_query($con,"insert into payment(sname,f_name,m_name,reg_no,mob,product,session,course,email,class,amount,trn,pdate,pstatus,pmode,gatetrn,paction,dept_roll,gen,dob,cat) values('$stu_name','$father','$mother','$reg_no','$mobile','$product','$session','$course','$email','$class','$amount','$trn','$pdate','$pstatus','$pmode','$gatetrn','$paction','$dept_roll','$gender','$dob','$category')");
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