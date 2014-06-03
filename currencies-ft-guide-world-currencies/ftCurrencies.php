<?php
//
//
//
//

if (empty($argv[1])) {
  echo 'Please enter input file' . "\n";
  echo 'php ftCurrencies.php path/to/data/file' . "\n";
  exit;
}

$inputFile = $argv[1];

if (!file_exists($inputFile)) {
  echo 'Input file does not exist! ' . $inputFile . "\n";
  exit;
}


$fileContents = file_get_contents($inputFile);
$fileContents = explode("\n", $fileContents);

$fileContents = array_map(function($l){
    return trim($l);
  },
  $fileContents
);

echo "Reading input file $inputFile\n";

// print_r($fileContents);
$output = array(
  array('Currency', 'GBP STG', 'GBP Week Change', 'US Dollar', 'US Dollar Week Change', 'Euro', 'Euro Week Change', 'Yen', 'Yen Week Change'),
);

$data = array();
$lastLine = array();

echo "Processing...\n";

foreach ($fileContents as $row) {

  if ($row && !preg_match('#STG|Change|GUIDE|:|Locking Rates|Reuters|\w \d+$#', $row)) {
    // print_r($row);

    $bits = explode(' ', $row);

    $currency = array();
    $line = array($currency);
    foreach ($bits as $b) {
      if (preg_match('#[a-zA-Z\(\)]#', $b)) {
        $line[0][] = $b;
      }
      else if (strpos($b, '&') !== FALSE) {
        $line[0][] = $b;
      }
      else {
        $line[] = $b;
      }
    }

    $line[0] = implode(' ', $line[0]);

    if (preg_match('#SDR#', $line[0])) {
      $lastLine[] = $line;
    }
    else {
      $data[$line[0]] = $line;
    }

  }
}

ksort($data);

$output = array_merge($output, array_values($data), $lastLine);

// print_r($output);

$outputFile = str_replace('.csv', '-cleaned.csv', $inputFile);

echo "Writing. $outputFile \n";

$fp = fopen($outputFile, 'w');

foreach ($output as $row) {
  // really annoying that fputcsv won't enclose all fields within enclosures!
  // http://stackoverflow.com/questions/1800675/write-csv-to-file-without-enclosures-in-php
  fputcsv($fp, $row, ',', '"');
}
fclose($fp);

echo "FINISH\n";





