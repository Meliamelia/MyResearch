<?PHP if ( ! defined('BASEPATH')) exit('No direct script access allowed');
class Cpanel extends CI_Controller{
	function Cpanel(){ 
		parent::__construct();
		$this->load->helper(array('form', 'url'));
		$this->load->library('Excel/Classes/PHPExcel');	
	}
	
	function index(){
		
	}	
	
	function DownloadExcell($id){
		$data = $this->Mpanel->cek_panel($id);
		$urlimg = $this->input->post("urlimg");
		$detail_urlimg = $this->input->post("detail_urlimg");
		$panelname = $data[0]['panelName'];
		
		//download panel
		if($urlimg){
		$objPHPExcel = new PHPExcel();  
        $objPHPExcel->setActiveSheetIndex(0)->setCellValue('B1', 'Panel')
											->setCellValue('C1', $panelname);
        
        $objPHPExcel->getActiveSheet()->setTitle($panelname);
		$gdImage = imagecreatefromjpeg($urlimg);
		$objDrawing = new PHPExcel_Worksheet_MemoryDrawing();
		$objDrawing->setName('Sample image');
		$objDrawing->setDescription('Sample image');
		$objDrawing->setImageResource($gdImage);
		$objDrawing->setRenderingFunction(PHPExcel_Worksheet_MemoryDrawing::RENDERING_JPEG);
		$objDrawing->setMimeType(PHPExcel_Worksheet_MemoryDrawing::MIMETYPE_DEFAULT);
		$objDrawing->setHeight(150);
		$objDrawing->setCoordinates('A3');
		$objDrawing->setWorksheet($objPHPExcel->getActiveSheet());
		$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
		
		
		//readfile ($objPHPExcel);
		header("Last-Modified: " . gmdate("D, d M Y H:i:s") . " GMT");
		header("Cache-Control: no-store, no-cache, must-revalidate");
		header("Cache-Control: post-check=0, pre-check=0", false);
		header("Pragma: no-cache");
		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		header("Content-Disposition: attachment;filename=".$panelname.".xlsx");
		$objWriter->save("php://output");
		}else{
		
		//download detail panel
		$objPHPExcel = new PHPExcel();  
        $objPHPExcel->setActiveSheetIndex(0)->setCellValue('B1', 'Detail')
											->setCellValue('C1', 'Panel')
											->setCellValue('D1', $panelname);
        
        $objPHPExcel->getActiveSheet()->setTitle($panelname);
		$gdImage = imagecreatefromjpeg($detail_urlimg);
		$objDrawing = new PHPExcel_Worksheet_MemoryDrawing();
		$objDrawing->setName('Sample image');
		$objDrawing->setDescription('Sample image');
		$objDrawing->setImageResource($gdImage);
		$objDrawing->setRenderingFunction(PHPExcel_Worksheet_MemoryDrawing::RENDERING_JPEG);
		$objDrawing->setMimeType(PHPExcel_Worksheet_MemoryDrawing::MIMETYPE_DEFAULT);
		$objDrawing->setHeight(150);
		$objDrawing->setCoordinates('A3');
		$objDrawing->setWorksheet($objPHPExcel->getActiveSheet());
		$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
		$objWriter->setIncludeCharts(TRUE);
		
		//readfile ($objPHPExcel);
		header("Last-Modified: " . gmdate("D, d M Y H:i:s") . " GMT");
		header("Cache-Control: no-store, no-cache, must-revalidate");
		header("Cache-Control: post-check=0, pre-check=0", false);
		header("Pragma: no-cache");
		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		header("Content-Disposition: attachment;filename="."Detail ".$panelname.".xlsx");
		$objWriter->save("php://output");
		}
        }
			
}

/* ---------------------------------------javascript--------------------------
$('.detail_download').click(function() {   
	var getchid = this.id;
	var judul = this.name;
    html2canvas($("#detail_section"+this.id), {
        onrendered: function(canvas) {    
            var imgData = canvas.toDataURL('image/jpeg', 1.0);  
            var doc = new jsPDF('p','mm');
			doc.text(judul,80,28);
            doc.addImage(imgData, 'JPEG', 10, 30);
		
			$("#detail_optdown").modal("show");
			$("#detail_topdf").click(function() {
				$("#detail_optdown").modal("hide");
				doc.save(judul+'.pdf'); 
			});
			$("#detail_toexcel").click(function() {
				$("#detail_optdown").modal("hide");
					$("#detail_urlimg").val(imgData);
					$("#detail_formtoexcel").attr('action', "<?PHP echo base_url(); ?>Cpanel/downloadExcell/"+getchid);
					$("#detail_formtoexcel").submit();
					
			});
        }
    });
}); 
---------------------------------------------------------------------------- */