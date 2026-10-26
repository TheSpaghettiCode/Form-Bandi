<?php
$file_path = 'data/output/compilazioni_docenti.xlsx';
$file_name = basename($file_path);
$file_exists = file_exists($file_path);

$db_path = __DIR__ . '/data/database.json';
$backup_dir = __DIR__ . '/data/backup';
if (!is_dir($backup_dir)) {
    mkdir($backup_dir, 0777, true);
}
$backup_excel = $backup_dir . '/backup_compilazioni.xlsx';
$backup_db = $backup_dir . '/backup_database.json';
$backup_exists = file_exists($backup_excel);

$message = '';
if (isset($_GET['action'])) {
    $action = $_GET['action'];
    
    if ($action === 'backup') {
        if ($file_exists) copy($file_path, $backup_excel);
        if (file_exists($db_path)) copy($db_path, $backup_db);
        $backup_exists = true;
        $message = '<div class="msg success">Backup di sicurezza creato con successo!</div>';
    }
    
    if ($action === 'restore') {
        if ($backup_exists) copy($backup_excel, $file_path);
        if (file_exists($backup_db)) copy($backup_db, $db_path);
        $file_exists = true;
        $message = '<div class="msg success">Dati ripristinati correttamente dall\'ultimo backup.</div>';
    }

    if ($action === 'delete') {
        if ($file_exists) copy($file_path, $backup_excel);
        if (file_exists($db_path)) copy($db_path, $backup_db);
        $backup_exists = true;

        if ($file_exists) unlink($file_path);
        if (file_exists($db_path)) unlink($db_path);
        $file_exists = false;
        $message = '<div class="msg success">I dati sono stati eliminati correttamente. È stato generato un <b>backup automatico</b> di sicurezza prima dell\'eliminazione.</div>';
    }

    if ($action === 'delete_backup') {
        if (file_exists($backup_excel)) unlink($backup_excel);
        if (file_exists($backup_db)) unlink($backup_db);
        $backup_exists = false;
        $message = '<div class="msg success">I file di backup sono stati eliminati definitivamente.</div>';
    }

    if ($action === 'sync') {
        require_once __DIR__ . '/vendor/autoload.php';
        if ($file_exists) {
            try {
                $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($file_path);
                $sheet = $spreadsheet->getSheetByName('Riepilogo Budget Attivita');
                
                if ($sheet) {
                    $highestCol = $sheet->getHighestDataColumn();
                    $highestColIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestCol);
                    
                    $colCorso = 'D'; $colNome = 'B'; $colCognome = 'C'; $colNote = 'M';
                    for ($col = 1; $col <= $highestColIndex; $col++) {
                        $letter = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col);
                        $headerVal = trim((string)$sheet->getCell($letter . '1')->getValue());
                        if ($headerVal === 'Corso Selezionato') $colCorso = $letter;
                        if ($headerVal === 'Nome Compilatore') $colNome = $letter;
                        if ($headerVal === 'Cognome Compilatore') $colCognome = $letter;
                        if ($headerVal === 'Note Generali') $colNote = $letter;
                    }
                    
                    $db = file_exists($db_path) ? (json_decode(file_get_contents($db_path), true) ?: []) : [];
                    
                    $corsiInExcel = [];
                    $corsiAggiunti = 0;
                    $highestRow = $sheet->getHighestDataRow();
                    
                    for ($r = 2; $r <= $highestRow; $r++) {
                        $corso = trim((string)$sheet->getCell($colCorso . $r)->getValue());
                        if (!empty($corso)) {
                            $corsiInExcel[] = $corso;
                            if (!isset($db[$corso])) {
                                $db[$corso] = [
                                    'nome' => trim((string)$sheet->getCell($colNome . $r)->getValue()),
                                    'cognome' => trim((string)$sheet->getCell($colCognome . $r)->getValue()),
                                    'note' => trim((string)$sheet->getCell($colNote . $r)->getValue())
                                ];
                                $corsiAggiunti++;
                            }
                        }
                    }
                    
                    $corsiRimossi = 0;
                    foreach ($db as $key => $val) {
                        if (!in_array($key, $corsiInExcel)) {
                            unset($db[$key]);
                            $corsiRimossi++;
                        }
                    }
                    
                    file_put_contents($db_path, json_encode($db, JSON_PRETTY_PRINT));
                    $message = '<div class="msg info">Sincronizzazione completata! ' . $corsiAggiunti . ' nuovi corsi aggiunti, ' . $corsiRimossi . ' corsi eliminati dal database.</div>';
                } else {
                    $message = '<div class="msg warning">Foglio di riepilogo non trovato nell\'Excel.</div>';
                }
            } catch (Exception $e) {
                $message = '<div class="msg error">Errore durante la sincronizzazione: ' . $e->getMessage() . '</div>';
            }
        } else {
            file_put_contents($db_path, json_encode([], JSON_PRETTY_PRINT));
            $message = '<div class="msg info">Sincronizzazione completata! Non essendoci alcun file Excel, il database JSON è stato svuotato.</div>';
        }
    }
}
?>
<!DOCTYPE html>
<html lang="it">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Pannello Amministrazione Dati</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            display: flex;
            justify-content: center;
            align-items: center;
            min-height: 100vh;
            margin: 0;
            background-color: #f4f6f8;
            padding: 20px;
        }
        .card {
            background: #fff;
            border-radius: 8px;
            box-shadow: 0 2px 12px rgba(0,0,0,0.15);
            padding: 30px 40px;
            max-width: 500px;
            width: 100%;
        }
        h1 {
            font-size: 1.5rem;
            color: #333;
            margin-top: 0;
            text-align: center;
            margin-bottom: 25px;
        }
        .section-box {
            padding: 20px;
            border-radius: 8px;
            margin-bottom: 20px;
            text-align: center;
        }
        .section-download {
            border: 2px solid #1a73e8;
            background-color: #f1f7fe;
        }
        .section-sync {
            border: 2px solid #fd7e14;
            background-color: #fff6f0;
        }
        .section-danger {
            border: 2px solid #dc3545;
            background-color: #fcf1f2;
        }
        h3 {
            margin-top: 0;
            margin-bottom: 15px;
            font-size: 1.1rem;
        }
        .section-download h3 { color: #1a73e8; }
        .section-sync h3 { color: #fd7e14; }
        .section-danger h3 { color: #dc3545; }
        
        .btn {
            display: inline-block;
            color: #fff;
            text-decoration: none;
            padding: 10px 20px;
            border-radius: 5px;
            font-size: 0.95rem;
            transition: background-color 0.2s;
            margin: 5px;
            border: none;
            cursor: pointer;
        }
        .btn-download { background-color: #1a73e8; }
        .btn-download:hover { background-color: #1558b0; }
        .btn-sync { background-color: #fd7e14; }
        .btn-sync:hover { background-color: #e8590c; }
        .btn-backup { background-color: #28a745; }
        .btn-backup:hover { background-color: #218838; }
        .btn-restore { background-color: #6c757d; }
        .btn-restore:hover { background-color: #5a6268; }
        .btn-delete { background-color: #dc3545; }
        .btn-delete:hover { background-color: #c82333; }
        .btn-delete-backup { background-color: #f06363ff; }
        .btn-delete-backup:hover { background-color: #ef4444; }
        
        .file-info {
            margin-top: 15px;
            font-size: 0.85rem;
            color: #666;
        }
        .msg {
            padding: 12px;
            border-radius: 5px;
            margin-bottom: 20px;
            font-size: 0.95rem;
            text-align: center;
        }
        .msg.success { background-color: #d4edda; color: #155724; }
        .msg.info { background-color: #cce5ff; color: #004085; }
        .msg.warning { background-color: #fff3cd; color: #856404; }
        .msg.error { background-color: #f8d7da; color: #721c24; }
    </style>
</head>
<body>
    <div class="card">
        <h1>📊 Pannello Amministrazione Dati</h1>

        <?= $message ?>

        <!-- SEZIONE DOWNLOAD -->
        <div class="section-box section-download">
            <h3>📥 Area Download</h3>
            <?php if ($file_exists): ?>
                <p style="font-size: 0.9rem; color: #555; margin-bottom: 15px;">Scarica l'ultimo file Excel generato dalle compilazioni.</p>
                <a class="btn btn-download" href="?action=download">⬇ Scarica <?= htmlspecialchars($file_name) ?></a>
                <div class="file-info">
                    Dimensione: <?= round(filesize($file_path) / 1024, 1) ?> KB &nbsp;|&nbsp; Modifica: <?= date('d/m/Y H:i', filemtime($file_path)) ?>
                </div>
            <?php else: ?>
                <p style="color: #c0392b; font-weight: bold; margin: 0;">⚠ Nessun file Excel trovato</p>
            <?php endif; ?>
        </div>
        
        <!-- SEZIONE SYNC E BACKUP -->
        <div class="section-box section-sync">
            <h3>🔄 Gestione e Sincronizzazione</h3>
            <p style="font-size: 0.9rem; color: #555; margin-bottom: 15px;">
                Mantieni il database JSON allineato con eventuali modifiche manuali fatte nell'Excel, o salva/ripristina lo stato attuale.
            </p>
            <a class="btn btn-sync" href="?action=sync" onclick="return confirm('Sincronizzare il database? I corsi aggiunti manualmente nell\'Excel verranno bloccati nel form, e quelli rimossi verranno sbloccati.');">
                🔄 Sincronizza Database
            </a>
            <div style="margin-top: 10px;">
                <?php if ($file_exists): ?>
                    <a class="btn btn-backup" href="?action=backup" onclick="return confirm('Vuoi creare un backup di sicurezza dei dati attuali (Excel e JSON)? Sovrascriverà l\'ultimo backup se presente.');">
                        💾 Crea Backup
                    </a>
                <?php endif; ?>
                
                <?php if ($backup_exists): ?>
                    <a class="btn btn-restore" href="?action=restore" onclick="return confirm('ATTENZIONE: Stai per ripristinare i dati dall\'ultimo backup. L\'Excel e il JSON attuali verranno SOVRASCRITTI dal backup. Procedere?');">
                        🔙 Ripristina Backup
                    </a>
                <?php endif; ?>
            </div>
            
            <?php if ($backup_exists): ?>
            <div class="file-info" style="font-size: 0.8rem; margin-top: 15px;">
                Ultimo backup: <?= date('d/m/Y H:i', filemtime($backup_excel)) ?>
            </div>
            <?php endif; ?>
        </div>

        <!-- SEZIONE PERICOLO -->
        <?php if ($file_exists || $backup_exists): ?>
        <div class="section-box section-danger">
            <h3>⚠️ Zona di Pericolo</h3>
            <p style="font-size: 0.9rem; color: #555; margin-bottom: 15px;">
                Elimina tutti i dati raccolti per azzerare il form o elimina i backup salvati.
            </p>
            
            <div style="display: flex; justify-content: center; gap: 10px; flex-wrap: wrap;">
                <?php if ($backup_exists): ?>
                <a class="btn btn-delete-backup" href="?action=delete_backup" onclick="return confirm('Sei sicuro di voler eliminare i file di backup? Questa operazione è irreversibile e non potrai più ripristinare i dati precedenti.');">
                    🧹 Elimina Backup
                </a>
                <?php endif; ?>
                
                <?php if ($file_exists): ?>
                <a class="btn btn-delete" href="?action=delete" onclick="return confirm('ATTENZIONE! Sei sicuro di voler eliminare TUTTI i dati raccolti? Verrà prima effettuato un backup di sicurezza in caso di errore. Questa operazione azzererà il file Excel.');">
                    🗑️ Elimina Tutti i Dati
                </a>
                <?php endif; ?>
            </div>
        </div>
        <?php endif; ?>

    </div>
</body>
</html>