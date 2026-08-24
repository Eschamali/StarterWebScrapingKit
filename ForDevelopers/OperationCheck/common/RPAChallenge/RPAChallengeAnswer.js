(function(allData) {
    // 待ち関数
    const sleep = ms => new Promise(res => setTimeout(res, ms));

    // Submitボタンも事前に取得しておく
    const submitBtn = document.querySelector('input[type="submit"]');

    // 1. Startボタンをクリック
    document.querySelector("body > app-root > div.body.row1.scroll-y > app-rpa1 > div > div.instructions.col.s3.m3.l3.uiColorSecondary > div:nth-child(7) > button").click();

    // 2. 10回戦のループ
    for (const rowData of allData) {
        const labels = Array.from(document.querySelectorAll('label'));

        // 3. 1ラウンド分のデータを入力
        for (const [key, value] of Object.entries(rowData)) {
            labels.find(l => l.innerText.trim() === key).parentElement.querySelector('input').value = value;
        }

        // 4. Submitをクリック
        submitBtn.click();
    }

    return "Mission Accomplished!";
})([
    {"First Name":"John","Last Name":"Smith","Role in Company":"Analyst","Address":"98 North Road", "Email":"jsmith@itsolutions.co.uk", "Phone Number":"40716543298", "Company Name":"IT Solutions"},
    {"First Name":"Jane","Last Name":"Dorsey","Role in Company":"Medical Engineer","Address":"11 Crown Street", "Email":"jdorsey@mc.com", "Phone Number":"40791345621", "Company Name":"MediCare"},
    {"First Name":"Albert","Last Name":"Kipling","Role in Company":"Accountant","Address":"22 Guild Street", "Email":"kipling@waterfront.com", "Phone Number":"40735416854", "Company Name":"Waterfront"},
    {"First Name":"Michael","Last Name":"Robertson","Role in Company":"IT Specialist","Address":"17 Farburn Terrace", "Email":"mrobertson@mc.com", "Phone Number":"40733652145", "Company Name":"MediCare"},
    {"First Name":"Doug","Last Name":"Derrick","Role in Company":"Analyst","Address":"99 Shire Oak Road", "Email":"dderrick@timepath.co.uk", "Phone Number":"40799885412", "Company Name":"Timepath Inc."},
    {"First Name":"Jessie","Last Name":"Marlowe","Role in Company":"Scientist","Address":"27 Cheshire Street", "Email":"jmarlowe@aperture.us", "Phone Number":"40733154268", "Company Name":"Aperture Inc."},
    {"First Name":"Stan","Last Name":"Hamm","Role in Company":"Advisor","Address":"10 Dam Road", "Email":"shamm@sugarwell.org", "Phone Number":"40712462257", "Company Name":"Sugarwell"},
    {"First Name":"Michelle","Last Name":"Norton","Role in Company":"Scientist","Address":"13 White Rabbit Street", "Email":"mnorton@aperture.us", "Phone Number":"40731254562", "Company Name":"Aperture Inc."},
    {"First Name":"Stacy","Last Name":"Shelby","Role in Company":"HR Manager","Address":"19 Pineapple Boulevard", "Email":"sshelby@techdev.com", "Phone Number":"40741785214", "Company Name":"TechDev"},
    {"First Name":"Lara","Last Name":"Palmer","Role in Company":"Programmer","Address":"87 Orange Street", "Email":"lpalmer@timepath.co.uk", "Phone Number":"40731653845", "Company Name":"Timepath Inc."}
]);
