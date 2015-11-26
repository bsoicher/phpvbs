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
			"{$function}", //post_name Permalink
			create_title($function) , // post_title
			create_body($function, $cat['basename'], $file), 
	);
	//print_r($line);

	$count++;
	
}
 

 

/***********************
 * create CSV file
 * *********************/
	$csv = new CsvHelper();
	$header = array(
				'function',
				'title',
				'body'
			); 

    $csv->addRow($header);

	foreach ($line as $data) {
		$csv->addRow($data);
	}
	$result = $csv->render(false, "UTF-8");
	$result = str_replace('"""', '"', $result);
	$r = file_put_contents(dirname(__FILE__) . "/wordpress_jptitle_and_jpbody.csv", $result);
 
 
 
 

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
	$title = "VBScript {$str} 関数";
	return $title;
}

/**
 *  create body strings
 */
function create_body($str, $cat, $code) {
	
	$linecount = 0; // 行数
	$subject = ''; // 引数や、戻り値、などの項目名
	$block = ''; //どのブロックを処理しているか
	
	$code_text = '';
	$pre_text = '';
	$hikisu_text = '';
	$return_text = '';
	$execute_text = '';
	
	$line = explode("\n",$code);

	foreach($line as $l) {
		$linecount++;
		$first = mb_substr($l ,0 ,1);
		if($first == "'") { // １文字目がコメントの場合に処理開始
			
			$line_text = mb_substr($l ,1);
			if(stristr($l, '========')) continue; //1行目と３行目は================とコメント行のため削除
			if($linecount == 2 ) {
				$pre_text = $line_text . "\n";
			}
			
// 			//項目名があるかどうかチェック
			$koumoku_line = false;
			if(stristr($l, '【引数】')) {
				$block = 'hikisu';
				$koumoku_line = true;
			} elseif(stristr($l, '【戻り値】')) {
				$block = 'return';
				$koumoku_line = true;
			} elseif(stristr($l, '【処理】')) {
				$block = 'execute';
				$koumoku_line = true;
			}
			
			if(!$koumoku_line) {
				switch ($block) {
				case 'hikisu':
					$hikisu_text .= $line_text . "\n";
					break;
				case 'return':
					$return_text .= $line_text . "\n";
					break;
				case 'execute':
					$execute_text .= $line_text . "\n";
					break;
				}
			}
			$koumoku_line = false;
			
			
		} else {
			$code_text .= $l. "\n";
		}
		
	}
	
$body = <<<EOM


{$pre_text}

<pre><code class="vbscript">
{$code_text}
</code></pre>

<ul>
	<li><a href="https://github.com/masayukiando/phpvbs/blob/master/functions/{$cat}/{$str}.vbs" target="_blank">GitHubで見る</a></li>
</ul>


<h3>引数</h3>
{$hikisu_text}

<h3>戻り値</h3>
{$return_text}

<h3>処理</h3>
{$execute_text}



<ul>
	<li><a href="https://github.com/masayukiando/phpvbs/edit/master/functions/{$cat}/{$str}.vbs" target="_blank">GitHubで編集</a></li>

</ul>


EOM;
	
return $body;
	
}