<?php
session_start();
if(isset($_SESSION['name'])){
    //$text = '';
    if(isset($_POST['text']))
    {
	$text = $_POST['text'];
    }   
        
    include 'Excel/reader.php';
    
    // initialize reader object
    $excel = new Spreadsheet_Excel_Reader();
    
    // read spreadsheet data
    $excel->read('book1.xls');    
    
    // iterate over spreadsheet cells and print as HTML table
    //echo $excel->sheets[0]['numRows'];
   // while($x<=$excel->sheets[0]['numRows']) {
      //echo "\t<tr>\n";
      $y=1;
      while($y<=$excel->sheets[0]['numCols']) {
        $cell = isset($excel->sheets[0]['cells'][$y][1]) ? $excel->sheets[0]['cells'][$y][1] : '';
        //echo "\t\t<td>$cell</td>\n";
        echo $excel->sheets[0]['cells'][$y][1];
        if($text === $cell)
        {
            $_SESSION['cell']=$excel->sheets[0]['cells'][$y][2];
        }
        
        //echo $cell;
        $y++;
      }  
      //echo "\t</tr>\n";
      //$x++;
  //  }
   
//        echo $text;
//	
	$fp = fopen("log.html", 'a');
	fwrite($fp, "<div class='msgln'>(".date("g:i A").") <b>".$_SESSION['name']."</b>: ".stripslashes(htmlspecialchars($text))."<br></div>");
        if(isset($_SESSION['cell'])){
            fwrite($fp, "<div class='msgln'>(".date("g:i A").") <b>Operator </b>: ".stripslashes(htmlspecialchars($_SESSION['cell']))."<br></div>");
            unset($_SESSION['cell']);
        }
        
	fclose($fp);
}
?>