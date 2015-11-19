<?php
/*
	*******************************
	** GS-Zone Web-side AO-Update
	** Copyright (c) 2013 - GS-Zone
	** Versión: 0.3
	*******************************
*/


global $up_dir;
$up_dir = './update/';
$ur_dir = './updater/';
$updater_file = 'GSZAU';
$s_exe = '.exe';


if(isset($_GET['updater'])) {
	// Verifica solo si el AOUpdater esta actualizado.
	if(is_file($ur_dir . $updater_file . $s_exe)){
		die($updater_file. "|" . filesize($ur_dir . $updater_file . $s_exe) . "|" . strtoupper(md5_file($ur_dir . $updater_file . $s_exe)) . chr(10));
	} else {
		die("ERROR 404"); // No se encuentra el programa de actualización
	}
}

die(getFileList($up_dir));

function getFileList($dir)
{
	$retorno = '';
    if(substr($dir, -1) != "/") $dir .= "";
    $d = @dir($dir) or die("ERROR 404"); // Directorio no encontrado
    while(false !== ($entry = $d->read())) {
      if($entry[0] == ".") continue; // Saltear archivos ocultos (inician con ".")
      if(is_dir("$dir$entry")) {
      	$retorno .= getFileList("$dir$entry/"); // Ingresamos al subdirectorio
      } elseif(is_readable("$dir$entry")) {
      	$retorno .= remove_up_dir("$dir$entry") . "|" . filesize("$dir$entry") . "|" . strtoupper(md5_file("$dir$entry")) . chr(10); // Archivo
      }
    }
    $d->close();

    return $retorno;
}

function remove_up_dir($string) {
	global $up_dir;
	$string = substr($string, (strlen($up_dir)-1), strlen($string) - (strlen($up_dir)-1));
	$string = str_replace('/', '\\', $string);
	return $string;
}

?>