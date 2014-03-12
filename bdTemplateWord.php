<?php
	require_once dirname(__FILE__) . '/PHPWord.php';

	class bdTemplateWord {
		private $objDocument = null,
				$strFileName = false
		;

		/**
		 * Magic Methods
		 */
		public function __construct($getFileTpl = false, $getStrCreator = false)
		{
			if (! is_file($getFileTpl)) throw new Exception("File '$getFileTpl' not found", 1);
			$objPHPWord = new PHPword();
			if ($getStrCreator){
				$objPHPWord->
					getProperties()->
						setCreator($getStrCreator)->
							setLastModifiedBy($getStrCreator)
				;
			}
			$this->objDocument = $objPHPWord->loadTemplate($getFileTpl);
			/**
			 * Set the filename
			 */
			$strFilename = basename($getFileTpl, '.'.pathinfo($getFileTpl, PATHINFO_EXTENSION));
			$this->setFileName($strFilename);
		}

		/**
		 * Setters
		 */	
			/**
			 * Filename (string) $getFileName filename without extension
			 */
			public function setFileName($getFileName = false)
			{
				$this->strFileName = $getFileName.time().'.docx';
			}

		/**
		 * Convenience methods
		 */
			/**
			 * Fill the document by array
			 * if it is an multidemensional array.. like 
			 * array(
			 * 	array('var' => 'value one'),
			 * 	array('var' => 'value value 2')
			 * )
			 * the document will be looped (multi page)
			 */
			public function fillTemplateByArray($getArray = array())
			{
				if (! $getArray) return;
				$i = 0;
				foreach ($getArray as $key => $val) {
					if (is_array($val)) {
						if ($i>0) $this->objDocument->AddPage();
						foreach ($val as $key2 => $val2) {
							$this->setValue($key2, $val2);							
						}
						$i++;
					} else {
						$this->setValue($key, $val);
					}

				}
			}
			public function setValue($getKey, $getValue = false)
			{
				/**
				 * filter $getValue for html 
				 */
					$getValue = str_replace('<br />', '<w:br/>' ,nl2br($getValue));
				/**
				 * Set The value 
				 */
				$this->objDocument->setValue($getKey, $getValue);
			}

			public function downloadFile()
			{
				$file = realpath(dirname(__FILE__)).'/'.$this->strFileName;
				$this->objDocument->save($file);
				header('Content-Type: application/vnd.ms-word');
				header('Content-Disposition: attachment;filename="'.$this->strFileName.'"');
				header('Cache-Control: max-age=0');
				// If you're serving to IE 9, then the following may be needed
				header('Cache-Control: max-age=1');

				// If you're serving to IE over SSL, then the following may be needed
				header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
				header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
				header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
				header ('Pragma: public'); // HTTP/1.0
				readfile($file);
				unlink($file);
				exit;
			}

		/**
		 * Static Methods
		 */
			/**
			 *	Render 
			 * 	@param (string) $getFileTpl the full filepath of the template .docx file
			 */
			public function render($getFileTpl = false, $getArrData = false, $getFileNameDownload = false, $getBoolDownload = true)
			{
				$objBdTemplateWord = new bdTemplateWord($getFileTpl);
				$objBdTemplateWord->fillTemplateByArray($getArrData);
				if ($getFileNameDownload) $objBdTemplateWord->setFileName($getFileNameDownload);
				if ($getBoolDownload) $objBdTemplateWord->downloadFile();
			}
		/**
		 * End
		 */
	};

?>