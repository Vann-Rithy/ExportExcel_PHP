<!DOCTYPE html>
<html>
<head>
    <title>Data Export to Excel</title>
</head>
<body>
    <h2>Export Data to Excel</h2>
    <form action="export.php" method="POST">
        <div class="form-group">
            <label for="companyName">Select Company:</label>
            <select class="form-control" id="companyName" name="company_name">
                <option value="All Company">All Company</option>
                <option value="Dynamic Network Logistics Co., Ltd.">Dynamic Network Logistics Co., Ltd.</option>
                <option value="KC Motors (Cambodia) Co., Ltd.">KC Motors (Cambodia) Co., Ltd.</option>
                <option value="HK Car Inspection (Cambodia) Co., LTD.">HK Car Inspection (Cambodia) Co., LTD.</option>
                <option value="SME News">SME News</option>
                <option value="Yuttaka Law Office">Yuttaka Law Office</option>
                <option value="Plexus Trust Co., Ltd.">Plexus Trust Co., Ltd.</option>
                <option value="Prosper Reit Co., Ltd.">Prosper Reit Co., Ltd.</option>
            </select>
        </div>
        <button type="submit" class="btn btn-primary">Export Data to Excel</button>
    </form>
</body>
</html>
