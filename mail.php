public function email()
  {
    $input = Input::get();
    $ids=$input['id'];
    $freq=$input['dynamic'][0];
    $hours=$input['dynamic'][1];
    $days=$input['dynamic'][2];
    $creative=$input['dynamic'][3];
    //data from db based on input id
    $data=Radio::getFullData($ids);
    $rate=[];
    $total=0;
    foreach ($data as $key => $value) {
      $value->frequency=$freq;
      $value->hours=$hours;
      $value->days=$days;
      $value->creative=$creative;
      $rate=$value->rate;
      $cost =  $freq * $rate * $hours * $days * ($creative/10);
      $total+=$freq * $rate * $hours * $days * ($creative/10);
      $value->cost=$cost;
    }
    //excel in mail
    $objPHPExcel = new PHPExcel();
    $objPHPExcel->setActiveSheetIndex(0)
                ->setCellValue('A1', 'Stations')
                ->setCellValue('B1', 'Language')
                ->setCellValue('C1', 'Card Rates(in Rs.)')
                //->setCellValue('D1','Reach')
                ->setCellValue('D1','Frequency')
                ->setCellValue('E1','No. of Hours')
                ->setCellValue('F1','Days')
                ->setCellValue('G1','Creative Length')
                ->setCellValue('H1', 'Cost(in Rs.)');

    $styleArray = array(
        'font'  => array(
            'bold'  => true,
            'color' => array('rgb' => '2E8B57'),
            'name'  => 'Verdana'
    ));   
    $i = 1;
    foreach ($data as $option) {
    ++$i;
    $objPHPExcel->setActiveSheetIndex(0)
       ->setCellValue("A{$i}", $option->station);

    $objPHPExcel->setActiveSheetIndex(0)
       ->setCellValue("B{$i}", $option->language);

    $objPHPExcel->setActiveSheetIndex(0)
       ->setCellValue("C{$i}",$option->rate);

    //$objPHPExcel->setActiveSheetIndex(0)
       //->setCellValue("D{$i}", $option->reach);

    $objPHPExcel->setActiveSheetIndex(0)
       ->setCellValue("D{$i}", $option->frequency);

    $objPHPExcel->setActiveSheetIndex(0)
       ->setCellValue("E{$i}", $option->hours);
       
    $objPHPExcel->setActiveSheetIndex(0)
       ->setCellValue("F{$i}", $option->days);
       
    $objPHPExcel->setActiveSheetIndex(0)
       ->setCellValue("G{$i}", $option->creative);         

    $objPHPExcel->setActiveSheetIndex(0)
       ->setCellValue("H{$i}",$option->cost);   

    }

    foreach(['A', 'B', 'C', 'D','E','F','G','H'] as $columnID) {
    $objPHPExcel->getActiveSheet()
       ->getColumnDimension($columnID)
       ->setAutoSize(true);
    }

    $objPHPExcel->getActiveSheet()->freezePane('A2');
    $objPHPExcel->getActiveSheet()->getStyle("A1:H1")->getFont()->setBold(true);
    $j = 0;
    for($j=0;$j<=$i;$j++) {
    $objPHPExcel->getActiveSheet()->getStyle("A{$j}:H{$j}")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    }


    // Set assistance text
    $contact_start_row = $i+2;
    $contact_end_row = $contact_start_row + 1;
    $style['font']['color'] = ['rgb' => '0F7F12'];
    //$objPHPExcel->getActiveSheet()->getRowDimension($contact_start_row)->setRowHeight(25);
    $objPHPExcel->setActiveSheetIndex(0)
        ->setCellValue("A{$contact_start_row}", "(Please Call Anjana on 0-76767-38042 for best rates & immediate assistance.)");
    //$objPHPExcel->setActiveSheetIndex(0)->getStyle("A{$contact_start_row}")->getAlignment()->setWrapText(true);
    // set font size
    $objPHPExcel->setActiveSheetIndex(0)->getStyle("A{$contact_start_row}")->applyFromArray($styleArray);
    // merge
    $objPHPExcel->setActiveSheetIndex(0)->mergeCells("A{$contact_start_row}:I{$contact_end_row}");
    // set text center
    $objPHPExcel->getActiveSheet()->getStyle("A{$contact_start_row}:I{$contact_end_row}")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    //$objPHPExcel->getActiveSheet()->getStyle("A{$contact_start_row}:I{$contact_end_row}")->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER); 
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
    $public_path = public_path() . '/uploads/radio/';
    $public_path .= "Radio Plan - " . date('Y-m-d H:i:s', strtotime('now')) . '.xls';
    $objWriter->save($public_path);
    //send mail with excel
    $address[] = $input['email'];
    $address=implode(",",$address);
    $address=explode(",",str_replace(" ","",$address));
    $from = 'help@themediaant.com';
    $message= $data;
    //calculating total cost and formatting in indian style
    function makecommaa($input){
            if(strlen($input)<=2)
            { return $input; }
            $length=substr($input,0,strlen($input)-2);
            $formatted_input = makecommaa($length).",".substr($input,-2);
            return $formatted_input;
            }
    function formatInIndianStyle($num){
            // This is my function
            $pos = strpos((string)$num, ".");
            //if ($pos === false) { $decimalpart="";}
            // else { $decimalpart= substr($num, $pos+1, 2); $num = substr($num,0,$pos); }
            if(strlen($num)>3 & strlen($num) <= 12){
                        $last3digits = substr($num, -3 );
                        $numexceptlastdigits = substr($num, 0, -3 );
                        $formatted = makecommaa($numexceptlastdigits);
                        $stringtoreturn = $formatted.",".$last3digits ;
            }elseif(strlen($num)<=3){
                        $stringtoreturn = $num ;
            }elseif(strlen($num)>12){
                        $stringtoreturn = number_format($num, 3);
            }
            if(substr($stringtoreturn,0,2)=="-,"){$stringtoreturn = "-".substr($stringtoreturn,2 );}
            return $stringtoreturn;
    }
    $total=formatInIndianStyle($total);
    //end of totalcost calculation

    //sending mails
    foreach ($address as $key => $value) {
      $address=trim($value);
      //validate if it is proper or not
      $pattern = "/^([a-zA-Z0-9])+([a-zA-Z\._-])*@([a-zA-Z0-9])+(\.)+([a-zA-Z0-9\._-]+)+ $/";
      if(preg_match($pattern,$address)){
      Mail::send('emails.message',['msg' => $message,'city'=>$input['city'],'total'=>$total],function($msg) use ($address,$from,$public_path){
       $msg->from($from,'The Media Ant');
       $msg->to($address)->subject('Radio Plan');
       $msg->cc('help@themediaant.com');
       $msg->bcc(array('priya.darshini@themediaant.com'));
       $msg->attach($public_path);
      });
      }
      else
      {
        continue;
      }
    }

    //save into email-store
    $save['email_id']=$input['email'];
    $save['path']=$public_path;
    $save['category']='Radio';
    EmailStore::addEmail($save);
    return Redirect::to('/radio');
  }
