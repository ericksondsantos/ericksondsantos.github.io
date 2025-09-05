<?php
$site_code = $_POST['site_code'];
$site_model = $_POST['site_model'];
$head_count = $_POST['head_count'];

$command = escapeshellcmd("python3 process.py $site_code $site_model $head_count");
$output = shell_exec($command);

echo $output;
?>