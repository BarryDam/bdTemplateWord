<?php
	require '../bdTemplateWord.php';
	bdTemplateWord::render(
		dirname(__FILE__) . '/testfile.docx',	
		array(
			'test'=>'dit is test :D ', 
			'dam' => 'mijn achternaam'
		)		
	);
?>