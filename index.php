<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>OLFU GSHEET CLASS RECORD UPDATE</title>
    <style>
        body { font-family: sans-serif; padding: 40px; background: #f4f4f9; color: #333; }
        .container { max-width: 500px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        h2 { margin-top: 0; color: #333; }
        .form-group { margin-bottom: 20px; }
        label { display: block; margin-bottom: 5px; font-weight: bold; }
        input[type="text"], input[type="file"] { width: 100%; padding: 10px; border: 1px solid #ddd; border-radius: 4px; box-sizing: border-box; }
        
        button { 
            width: 100%; padding: 12px; background-color: #28a745; 
            color: white; border: none; border-radius: 4px; 
            font-size: 16px; cursor: pointer; transition: background 0.3s;
        }
        button:hover { background-color: #218838; }
        button:disabled { background-color: #ccc; cursor: not-allowed; }

        .note { font-size: 0.85em; color: #666; margin-top: 15px; background: #fff3cd; padding: 10px; border-radius: 4px; }

        /* --- SPINNER STYLES --- */
        #loading-overlay {
            display: none;
            text-align: center;
            margin-top: 20px;
            padding: 20px;
            background: #f8f9fa;
            border-radius: 4px;
            border: 1px solid #e9ecef;
        }

        .spinner {
            border: 4px solid #f3f3f3;
            border-top: 4px solid #3498db;
            border-radius: 50%;
            width: 30px;
            height: 30px;
            animation: spin 1s linear infinite;
            margin: 0 auto 10px auto;
        }

        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }

        /* --- FOOTER STYLES --- */
        footer {
            text-align: center;
            margin-top: 40px;
            font-size: 0.85rem;
            color: #888;
        }
    </style>
</head>
<body>

<div class="container">
    <h2>OLFU GSHEET CLASS RECORD UPDATE TOOL</h2>
    
    <form action="sync_students.php" method="POST" enctype="multipart/form-data" id="syncForm">
        
        <div class="form-group">
            <label for="sheet_name">Target Sheet Name (Tab Name):</label>
            <input type="text" name="sheet_name" id="sheet_name" placeholder="e.g. Sheet1" required>
        </div>

        <div class="form-group">
            <label for="excel_file">Upload Excel List (.xlsx):</label>
            <input type="file" name="excel_file" id="excel_file" accept=".xlsx, .xls" required>
        </div>

        <button type="submit" id="submitBtn">Start Sync Process</button>
    </form>

    <div id="loading-overlay">
        <div class="spinner"></div>
        <div id="loading-text">Analyzing file and syncing with Google Sheets...<br><small>Do not close this window.</small></div>
    </div>
    
    <div class="note">
        <strong>Note:</strong> This will insert rows for new students. Existing grades/attendance for current students will not be touched.
    </div>
</div>

<footer>
    Developed by AJ Retuya &copy; 2026 version 1.1
</footer>

<script>
    const form = document.getElementById('syncForm');
    const loader = document.getElementById('loading-overlay');
    const btn = document.getElementById('submitBtn');

    form.addEventListener('submit', function() {
        loader.style.display = 'block';
        btn.textContent = 'Processing...';
        btn.disabled = true;
    });
</script>

</body>
</html>