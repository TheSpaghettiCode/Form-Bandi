<?php
// FASE 3: SALVATAGGIO DATI
header('Content-Type: application/json');

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    echo json_encode(['success' => false, 'message' => 'Metodo non consentito.']);
    exit;
}

if (!file_exists(dirname(__DIR__) . '/vendor/autoload.php')) {
    echo json_encode(['success' => false, 'message' => 'Libreria PhpSpreadsheet non trovata. Esegui composer install.']);
    exit;
}

require_once dirname(__DIR__) . '/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

function sanitize($input) {
    if (is_null($input)) return '';
    return htmlspecialchars(strip_tags(trim($input)), ENT_QUOTES, 'UTF-8');
}

$postData = $_POST;
if (empty($postData)) {
    $raw = file_get_contents('php://input');
    $decoded = json_decode($raw, true);
    if (is_array($decoded)) {
        $postData = $decoded;
    }
}

// Estrai e sanitizza i dati fissi
$nome = sanitize($postData['nome'] ?? '');
$cognome = sanitize($postData['cognome'] ?? '');
$corso = sanitize($postData['corso'] ?? '');
$ssd = sanitize($postData['ssd'] ?? '');
$laurea = sanitize($postData['laurea'] ?? '');
$anno = sanitize($postData['anno'] ?? '');
$semestre = sanitize($postData['semestre'] ?? '');
$docente = sanitize($postData['docente'] ?? '');
$budget_iniziale = (float)($postData['budget_iniziale'] ?? 0);
$note = sanitize($postData['note'] ?? '');

// Estrazione campi attività dinamici da Budget_Architettura.xlsx
$fileTariffe = dirname(__DIR__) . '/data/input/Budget_Architettura.xlsx';
$titoliAttivita = [];
$tariffe = [];

if (file_exists($fileTariffe)) {
    $spreadsheetBudget = IOFactory::load($fileTariffe);
    $sheetBudget = $spreadsheetBudget->getActiveSheet();
    $highestRowBudget = $sheetBudget->getHighestDataRow();
    for ($row = 2; $row <= $highestRowBudget; $row++) {
        $nomeAtt = trim((string)$sheetBudget->getCell('A' . $row)->getValue());
        $costoStr = (string)$sheetBudget->getCell('B' . $row)->getCalculatedValue();
        $costoVal = (float)filter_var($costoStr, FILTER_SANITIZE_NUMBER_FLOAT, FILTER_FLAG_ALLOW_FRACTION);
        
        $conSpesaStr = trim((string)$sheetBudget->getCell('C' . $row)->getCalculatedValue());
        $conSpesa = ($conSpesaStr === '1' || strtolower($conSpesaStr) === 'true' || strtolower($conSpesaStr) === 'si');
        
        if (!empty($nomeAtt)) {
            $titoliAttivita[] = $nomeAtt;
            $tariffe[$nomeAtt] = ['costo' => $costoVal, 'con_spesa' => $conSpesa];
        }
    }
}

$allTariffeZero = true;
foreach ($tariffe as $val) {
    if (isset($val['costo']) && $val['costo'] > 0) $allTariffeZero = false;
}

if (empty($tariffe) || $allTariffeZero) {
    $titoliAttivita = ['Supporto contratti', 'Supporto studenti', 'Conferenze', 'Tutorato studenti', 'Tutorato dottorandi'];
    $tariffe = [
        'Supporto contratti' => ['costo' => 32.00, 'con_spesa' => true],
        'Supporto studenti' => ['costo' => 20.00, 'con_spesa' => true],
        'Conferenze' => ['costo' => 0, 'con_spesa' => false],
        'Tutorato studenti' => ['costo' => 18.00, 'con_spesa' => true],
        'Tutorato dottorandi' => ['costo' => 22.00, 'con_spesa' => true]
    ];
}

$datiAttivita = [];
foreach ($titoliAttivita as $titolo) {
    $inputName = strtolower(str_replace(' ', '_', $titolo));
    $dettagliRaw = $postData["{$inputName}_dettagli"] ?? '[]';
    
    json_decode($dettagliRaw);
    if (json_last_error() !== JSON_ERROR_NONE) {
        $dettagliRaw = '[]';
    }

    $datiAttivita[$titolo] = [
        'attiva' => isset($postData["{$inputName}_attiva"]) ? 1 : 0,
        'colloquio' => isset($postData["{$inputName}_colloquio"]) ? 1 : 0,
        'dettagli_json' => $dettagliRaw
    ];
}

$timestamp = date('d/m/Y H:i:s');
$dirPath = dirname(__DIR__) . '/data';
if (!is_dir($dirPath)) {
    if (!mkdir($dirPath, 0777, true)) {
        echo json_encode(['success' => false, 'message' => 'Impossibile creare la cartella data/.']);
        exit;
    }
}

// 1. Lettura / Aggiornamento JSON Database interno per il ripopolamento Frontend
$dbFile = $dirPath . '/database.json';
$db = [];
if (file_exists($dbFile)) {
    $dbData = file_get_contents($dbFile);
    if ($dbData) {
        $db = json_decode($dbData, true) ?: [];
    }
}

// Struttura base per questo corso nel DB
$corsoData = [
    'nome' => $nome,
    'cognome' => $cognome,
    'note' => $note
];

$richiesteFormattateArr = [];
$totaleRichiesto = 0;

foreach ($titoliAttivita as $titolo) {
    if ($datiAttivita[$titolo]['attiva']) {
        $inputName = strtolower(str_replace(' ', '_', $titolo));
        $corsoData["{$inputName}_attiva"] = true;
        $corsoData["{$inputName}_colloquio"] = $datiAttivita[$titolo]['colloquio'] == 1;
        
        $dettagliArr = json_decode($datiAttivita[$titolo]['dettagli_json'], true) ?: [];
        $corsoData["{$inputName}_sottovoci"] = $dettagliArr;
        
        $tariffa = $tariffe[$titolo]['costo'] ?? 0;
        $isPrice = $tariffe[$titolo]['con_spesa'] ?? false;
        
        $dettagliTesto = [];
        foreach ($dettagliArr as $det) {
            if ($isPrice) {
                $stud = (int)($det['studenti'] ?? 0);
                $ore = (int)($det['ore'] ?? 0);
                $costoSingolo = ($stud * $ore * $tariffa);
                $totaleRichiesto += $costoSingolo;
                $dettagliTesto[] = "$ore ore, $stud persone";
            } else {
                $costoSingolo = (float)($det['costo'] ?? 0);
                $totaleRichiesto += $costoSingolo;
                $dettagliTesto[] = "€" . number_format($costoSingolo, 2);
            }
        }
        
        if (!empty($dettagliTesto)) {
            $richiesteFormattateArr[] = "($titolo: " . implode(" + ", $dettagliTesto) . ")";
        }
    }
}

// Salva nel DB
$db[$corso] = $corsoData;
file_put_contents($dbFile, json_encode($db, JSON_PRETTY_PRINT));

$richiesteFormattateStr = implode(" ", $richiesteFormattateArr);

// 2. Creazione / Aggiornamento File Unico
$outDirPath = dirname(__DIR__) . '/data/output';
if (!is_dir($outDirPath)) {
    mkdir($outDirPath, 0777, true);
}
$mainFile = $outDirPath . '/compilazioni_docenti.xlsx';

try {
    if (file_exists($mainFile)) {
        $spreadsheet = IOFactory::load($mainFile);
    } else {
        $spreadsheet = new Spreadsheet();
    }

    // GESTIONE FOGLIO DI RIEPILOGO
    $sheetRiepTitle = 'Riepilogo Budget Attivita';
    $sheetRiep = $spreadsheet->getSheetByName($sheetRiepTitle);
    
    if (!$sheetRiep) {
        // Se è un file nuovo, usiamo il primo foglio esistente
        if ($spreadsheet->getSheetCount() == 1 && $spreadsheet->getSheet(0)->getTitle() == 'Worksheet') {
            $sheetRiep = $spreadsheet->getSheet(0);
        } else {
            $sheetRiep = $spreadsheet->createSheet();
        }
        $sheetRiep->setTitle($sheetRiepTitle);
        
        $headers = [
            'Data/Ora', 'Nome Compilatore', 'Cognome Compilatore', 'Corso Selezionato',
            'SSD', 'Laurea', 'Anno', 'Semestre',
            'Docente di Riferimento', 'Budget Previsto', 'Budget Richiesto Totale', 
            'Richieste (Attività)', 'Note Generali'
        ];
        
        $colIndex = 1;
        $sheetRiep->freezePane('E2');
        foreach ($headers as $index => $header) {
            $colLetter = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($colIndex);
            $sheetRiep->setCellValue($colLetter . '1', $header);
            $sheetRiep->getStyle($colLetter . '1')->getFont()->setBold(true);
            $sheetRiep->getColumnDimension($colLetter)->setAutoSize(true);
            $sheetRiep->getStyle($colLetter . '1')->getFill()
                ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                ->getStartColor()->setARGB('FF1F4E78');
            $sheetRiep->getStyle($colLetter . '1')->getFont()->getColor()->setARGB('FFFFFFFF');
            $colIndex++;
        }
    }

    // Trova colonna "Corso Selezionato" (dovrebbe essere D)
    $highestColRiep = $sheetRiep->getHighestDataColumn();
    $highestColIndexRiep = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColRiep);
    $colCorsoLetterRiep = 'D';
    for ($col = 1; $col <= $highestColIndexRiep; $col++) {
        $letter = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col);
        if (trim((string)$sheetRiep->getCell($letter . '1')->getValue()) === 'Corso Selezionato') {
            $colCorsoLetterRiep = $letter;
            break;
        }
    }

    $highestRowRiep = $sheetRiep->getHighestDataRow();
    // Cancella riga vecchia
    for ($r = $highestRowRiep; $r >= 2; $r--) {
        $corsoEsistente = trim((string)$sheetRiep->getCell($colCorsoLetterRiep . $r)->getValue());
        if ($corsoEsistente === trim($corso) && !empty($corsoEsistente)) {
            $sheetRiep->removeRow($r);
        }
    }

    $highestRowRiep = $sheetRiep->getHighestDataRow();
    $currentRowRiep = $highestRowRiep + 1;

    $sheetRiep->setCellValue("A$currentRowRiep", $timestamp);
    $sheetRiep->setCellValue("B$currentRowRiep", $nome);
    $sheetRiep->setCellValue("C$currentRowRiep", $cognome);
    $sheetRiep->setCellValue("D$currentRowRiep", $corso);
    $sheetRiep->setCellValue("E$currentRowRiep", $ssd);
    $sheetRiep->setCellValue("F$currentRowRiep", $laurea);
    $sheetRiep->setCellValue("G$currentRowRiep", $anno);
    $sheetRiep->setCellValue("H$currentRowRiep", $semestre);
    $sheetRiep->setCellValue("I$currentRowRiep", $docente);
    $sheetRiep->setCellValue("J$currentRowRiep", $budget_iniziale);
    $sheetRiep->setCellValue("K$currentRowRiep", $totaleRichiesto);
    $sheetRiep->setCellValue("L$currentRowRiep", $richiesteFormattateStr);
    $sheetRiep->setCellValue("M$currentRowRiep", $note);

    // GESTIONE FOGLI ATTIVITA'
    foreach ($titoliAttivita as $titolo) {
        $safeTitolo = substr(preg_replace('/[^A-Za-z0-9_\- ]/', '', $titolo), 0, 31);
        
        $sheetAtt = $spreadsheet->getSheetByName($safeTitolo);
        if (!$sheetAtt) {
            if (!$datiAttivita[$titolo]['attiva']) {
                continue;
            }
            $sheetAtt = $spreadsheet->createSheet();
            $sheetAtt->setTitle($safeTitolo);
            
            $isPriceTemplate = $tariffe[$titolo]['con_spesa'] ?? false;
            $dettagliArrTmp = json_decode($datiAttivita[$titolo]['dettagli_json'], true) ?: [];
            $hasRicerca = false;
            foreach ($dettagliArrTmp as $det) {
                if (isset($det['ricerca']) && trim($det['ricerca']) !== '') {
                    $hasRicerca = true;
                    break;
                }
            }
            if (strtolower(trim($titolo)) === 'supporto contratti') {
                $hasRicerca = true;
            }
            
            $headersAtt = [
                'Data/Ora', 'Nome Compilatore', 'Cognome Compilatore', 'Corso Selezionato',
                'SSD', 'Laurea', 'Anno', 'Semestre',
                'Docente di Riferimento', 'Nome Attività'
            ];
            
            if ($isPriceTemplate) {
                $headersAtt[] = 'Numero Persone';
                $headersAtt[] = 'Numero Ore';
                $headersAtt[] = 'Tariffa Oraria (€)';
            }
            
            $headersAtt[] = 'Costo Totale Richiesto (€)';
            $headersAtt[] = 'Colloquio';
            $headersAtt[] = 'Competenze/Requisiti';
            
            if ($hasRicerca) {
                $headersAtt[] = 'Titolo Ricerca';
            }
            
            $colIndex = 1;
            foreach ($headersAtt as $index => $header) {
                $colLetter = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($colIndex);
                $sheetAtt->setCellValue($colLetter . '1', $header);
                $sheetAtt->getStyle($colLetter . '1')->getFont()->setBold(true);
                $sheetAtt->getColumnDimension($colLetter)->setAutoSize(true);
                $sheetAtt->getStyle($colLetter . '1')->getFill()
                    ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                    ->getStartColor()->setARGB('FF1F4E78');
                $sheetAtt->getStyle($colLetter . '1')->getFont()->getColor()->setARGB('FFFFFFFF');
                $colIndex++;
            }
        }
        
        $highestColAtt = $sheetAtt->getHighestDataColumn();
        $highestColIndexAtt = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColAtt);
        $colCorsoLetterAtt = 'D';
        for ($col = 1; $col <= $highestColIndexAtt; $col++) {
            $letter = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col);
            if (trim((string)$sheetAtt->getCell($letter . '1')->getValue()) === 'Corso Selezionato') {
                $colCorsoLetterAtt = $letter;
                break;
            }
        }

        $highestRowAtt = $sheetAtt->getHighestDataRow();
        for ($r = $highestRowAtt; $r >= 2; $r--) {
            $corsoEsistente = trim((string)$sheetAtt->getCell($colCorsoLetterAtt . $r)->getValue());
            if ($corsoEsistente === trim($corso) && !empty($corsoEsistente)) {
                $sheetAtt->removeRow($r);
            }
        }

        if ($datiAttivita[$titolo]['attiva']) {
            $highestRowAtt = $sheetAtt->getHighestDataRow();
            $currentRowAtt = $highestRowAtt + 1;
            
            $dettagliArr = json_decode($datiAttivita[$titolo]['dettagli_json'], true) ?: [];
            $tariffa = $tariffe[$titolo]['costo'] ?? 0;
            $isPrice = $tariffe[$titolo]['con_spesa'] ?? false;
            
            $hasRicercaLoop = false;
            foreach ($dettagliArr as $det) {
                if (isset($det['ricerca']) && trim($det['ricerca']) !== '') {
                    $hasRicercaLoop = true;
                    break;
                }
            }
            if (strtolower(trim($titolo)) === 'supporto contratti') {
                $hasRicercaLoop = true;
            }
            
            foreach ($dettagliArr as $det) {
                $costoSingolo = 0;
                $stud = 0;
                $ore = 0;
                if ($isPrice) {
                    $stud = (int)($det['studenti'] ?? 0);
                    $ore = (int)($det['ore'] ?? 0);
                    $costoSingolo = ($stud * $ore * $tariffa);
                } else {
                    $costoSingolo = (float)($det['costo'] ?? 0);
                }
                
                $competenzeDettaglio = $det['competenze'] ?? '';
                $ricercaDettaglio = $det['ricerca'] ?? '';

                $col = 1;
                $sheetAtt->setCellValue(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col++) . $currentRowAtt, $timestamp);
                $sheetAtt->setCellValue(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col++) . $currentRowAtt, $nome);
                $sheetAtt->setCellValue(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col++) . $currentRowAtt, $cognome);
                $sheetAtt->setCellValue(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col++) . $currentRowAtt, $corso);
                $sheetAtt->setCellValue(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col++) . $currentRowAtt, $ssd);
                $sheetAtt->setCellValue(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col++) . $currentRowAtt, $laurea);
                $sheetAtt->setCellValue(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col++) . $currentRowAtt, $anno);
                $sheetAtt->setCellValue(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col++) . $currentRowAtt, $semestre);
                $sheetAtt->setCellValue(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col++) . $currentRowAtt, $docente);
                $sheetAtt->setCellValue(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col++) . $currentRowAtt, $titolo);
                
                if ($isPrice) {
                    $sheetAtt->setCellValue(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col++) . $currentRowAtt, $stud);
                    $sheetAtt->setCellValue(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col++) . $currentRowAtt, $ore);
                    $sheetAtt->setCellValue(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col++) . $currentRowAtt, $tariffa);
                }
                
                $sheetAtt->setCellValue(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col++) . $currentRowAtt, $costoSingolo);
                $sheetAtt->setCellValue(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col++) . $currentRowAtt, $datiAttivita[$titolo]['colloquio'] ? 'SI' : 'NO');
                $sheetAtt->setCellValue(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col++) . $currentRowAtt, $competenzeDettaglio);
                
                if ($hasRicercaLoop) {
                    $sheetAtt->setCellValue(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col++) . $currentRowAtt, $ricercaDettaglio);
                }
                
                $currentRowAtt++;
            }
        }
    }
    
    $spreadsheet->setActiveSheetIndex(0);
    $writer = new Xlsx($spreadsheet);
    $writer->save($mainFile);

    echo json_encode([
        'success' => true, 
        'message' => 'Richiesta salvata con successo in compilazioni_docenti.xlsx!'
    ]);

} catch (Exception $e) {
    echo json_encode([
        'success' => false, 
        'message' => 'Errore nel salvataggio su file Excel: ' . $e->getMessage()
    ]);
}
?>
