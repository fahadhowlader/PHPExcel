<?php
require_once 'Classes/PHPExcel.php';
require_once 'Classes/PHPExcel/IOFactory.php';

$servername = "rds.goama.co";
$username = "fahad";
$password = "@f@HggAd";
$dbname = "ggtpUserSvc";

// Create connection
$conn = new mysqli($servername, $username, $password, $dbname);
// Check connection
if ($conn->connect_error) {
  die("Connection failed: " . $conn->connect_error);
}

$sql = "select distinct(game) from tournaments_score;";
$result = $conn->query($sql);

if ($result->num_rows > 0) {
  // output data of each row
  /* Create new PHPExcel object*/
	$objPHPExcel = new PHPExcel();
	$i=0;
	while($row = $result->fetch_assoc()) {
		$query = 'select id, game, `timestamp`, participant_id, value as score, duration, format((value/duration),2) as max_score from tournaments_score 
		where game = "'.$row['game'].'" and duration > 0 and value > 0 order by max_score DESC limit 2000;';
		$res = $conn->query($query);
		/* Create a first sheet, representing sales data*/
		$objPHPExcel->setActiveSheetIndex($i);
		$objPHPExcel->getActiveSheet()->setCellValue('A1', 'ID');
		$objPHPExcel->getActiveSheet()->setCellValue('B1', 'participant_id');
		$objPHPExcel->getActiveSheet()->setCellValue('C1', 'score');
		$objPHPExcel->getActiveSheet()->setCellValue('D1', 'duration');
		$objPHPExcel->getActiveSheet()->setCellValue('E1', 'max_score');
		$j=2;
		while($game_row = $res->fetch_assoc()) {
			$objPHPExcel->getActiveSheet()->setCellValue("A$j",$game_row["id"]);
			$objPHPExcel->getActiveSheet()->setCellValue("B$j",$game_row["participant_id"]);
			$objPHPExcel->getActiveSheet()->setCellValue("C$j",$game_row["score"]);
			$objPHPExcel->getActiveSheet()->setCellValue("D$j",$game_row["duration"]);
			$objPHPExcel->getActiveSheet()->setCellValue("E$j",$game_row["max_score"]);
		$j++;
		}
		/*Rename sheet*/
		$objPHPExcel->getActiveSheet()->setTitle($row['game']);
		$i++;
	}
	/* Redirect output to a client’s web browser (Excel5)*/
	header('Content-Type: application/vnd.ms-excel');
	header('Content-Disposition: attachment;filename="allGameReports.xls"');
	header('Cache-Control: max-age=0');
	$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
	$objWriter->save('php://output');
} else {
  echo "0 results";
}
$conn->close();