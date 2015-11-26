<?php
error_reporting(E_ALL ^ E_NOTICE);
setlocale(LC_ALL,'ja_JP');
date_default_timezone_set('Asia/Tokyo');

require_once( dirname(__FILE__) . '/csv.php');


//search directory 
$dir = dirname(__FILE__) . '/../functions/';

$result = list_files($dir);

$count = 1;

foreach($result as $r) {
	//if( !stristr($r, 'preg_quote') ) continue;
	
	
	$path_parts = pathinfo($r);
	$cat = pathinfo($path_parts['dirname']);
	$file = file_get_contents($r);
	$file = mb_convert_encoding($file, "UTF-8", "SJIS");

	$function = str_replace('.vbs', '', $path_parts['basename']);
	
	$line[] = array(
			'',//post_id
			"vbscript-{$function}-function", //post_name Permalink
			'admin', //post_author
			date("Y/m/d", strtotime("-{$count} day"  )), //post_date
			'post', //post_type
			'publish', //post_status
			create_title($function) , // post_title
			create_body($function, $cat['basename'], $file), //post_content
			$cat['basename'], //post_category
			'', //post_tags
	);
	//print_r($line);

	$count++;
	
}
 

 

/***********************
 * create CSV file
 * *********************/
	$csv = new CsvHelper();
	$header = array(
				'post_id',
				'post_name',
				'post_author',
				'post_date',
				'post_type',
				'post_status',
				'post_title',
				'post_content',
				'post_category',
				'post_tags'
			); 

    $csv->addRow($header);

	foreach ($line as $data) {
		$csv->addRow($data);
	}
	$result = $csv->render(false, "UTF-8");
	$result = str_replace('"""', '"', $result);
	$r = file_put_contents(dirname(__FILE__) . "/wordpress_csvimporter.csv", $result);
 
 
 
 

 /**
 *  directory search multi
 */
function list_files($dir){
    $list = array();
    $files = scandir($dir);
    foreach($files as $file){
        if($file == '.' || $file == '..' || $file == '.DS_Store'){
            continue;
        } else if (is_file($dir . $file)){
            $list[] = $dir . $file;
        } else if( is_dir($dir . $file) ) {
            //$list[] = $dir;
            $list = array_merge($list, list_files($dir . $file . DIRECTORY_SEPARATOR));
        }
    }
    return $list;
}

/**
 *  create title strings
 */
function create_title($str) {
	$title = "VBScript {$str} function";
	return $title;
}

/**
 *  create body strings
 */
function create_body($str, $cat, $code) {
	
	$pre = '';
	$line = explode("\n",$code);
	foreach($line as $l) {
		$first = mb_substr($l ,0 ,1);
		if($first == "'") continue;
		$pre .= $l. "\n";
	}
	
$body = <<<EOM
A VBScript equivalent of PHP’s {$str}

<pre><code class="vbscript">
{$pre}
</code></pre>

<ul>
	<li><a href="https://github.com/masayukiando/phpvbs/blob/master/functions/{$cat}/{$str}.vbs" target="_blank">Raw function on GitHub</a></li>
</ul>


Please also note that php.vbs offers community built functions and goes by the <a href="https://medium.com/@ienjoy/mcdonalds-theory-9216e1c9da7d" target="_blank">McDonald’s Theory</a>. We’ll put online functions that are far from perfect, in the hopes to spark better contributions. Do you have one? Then please just:

<ul>
	<li><a href="https://github.com/masayukiando/phpvbs/edit/master/functions/{$cat}/{$str}.vbs" target="_blank">Edit on GitHub</a></li>

</ul>


EOM;
	
return $body;
	
}