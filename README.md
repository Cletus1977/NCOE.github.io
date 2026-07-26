# NCOE.github.io
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>NCOE Prospective Teachers Digital Readiness Survey</title>
    <!-- SheetJS for Excel generation -->
    <script src="https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js"></script>
    <!-- QRCode.js library -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"></script>
    <style>
        body { font-family: Arial, sans-serif; background: #f0f4f8; margin: 20px; }
        .container { max-width: 1100px; margin: auto; background: white; padding: 30px; border-radius: 12px; box-shadow: 0 0 20px rgba(0,0,0,0.1); }
        h1 { text-align: center; color: #1e3a8a; }
        h2 { color: #1e40af; border-bottom: 3px solid #bfdbfe; padding-bottom: 10px; margin-top: 35px; }
        .question { margin: 18px 0; padding: 16px; background: #f8fafc; border-radius: 8px; border-left: 6px solid #3b82f6; }
        label { font-weight: bold; display: block; margin-bottom: 8px; }
        .likert { display: flex; flex-wrap: wrap; gap: 12px; margin-top: 8px; }
        .likert label { background: #e6f0ff; padding: 8px 12px; border-radius: 6px; cursor: pointer; font-weight: normal; }
        .likert input[type="radio"] { margin-right: 4px; }
        input[type="text"], input[type="number"], select, textarea {
            width: 100%; max-width: 600px; padding: 8px 10px; border: 1px solid #cbd5e1;
            border-radius: 6px; font-size: 15px; margin-top: 4px;
        }
        .btn-row { display: flex; justify-content: center; margin: 50px auto 20px; }
        .btn-submit { background: #1e40af; color: white; padding: 16px 45px; font-size: 18px; border: none; border-radius: 10px; cursor: pointer; }
        .btn-submit:hover { background: #1e3a8a; }
        .error { color: red; display: none; font-weight: normal; margin-left: 8px; }
        .note { text-align:center; font-size:15px; color:#555; margin-top:10px; }
        .success-banner {
            display:none; background:#d1fae5; border:2px solid #059669; color:#065f46;
            padding:14px 20px; border-radius:8px; text-align:center; font-size:16px;
            margin-top: 20px;
        }
        .reverse-note { font-size: 13px; color: #e11d48; margin-top: 4px; }
        .instruction-box {
            background: #eff6ff; border: 2px solid #3b82f6; border-radius: 8px;
            padding: 16px; margin: 20px 0; font-size: 15px; color: #1e40af;
        }
        /* QR Code Section */
        .qr-section {
            text-align: center;
            background: #f8fafc;
            border: 2px dashed #3b82f6;
            border-radius: 12px;
            padding: 25px;
            margin: 25px 0 35px 0;
        }
        .qr-section h3 {
            color: #1e40af;
            margin-top: 0;
            margin-bottom: 8px;
        }
        .qr-section p {
            color: #555;
            margin: 8px 0 18px 0;
            font-size: 15px;
        }
        #qrcode {
            display: inline-block;
            padding: 12px;
            background: white;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.15);
        }
        #qrcode img, #qrcode canvas {
            display: block;
        }
        .qr-url {
            margin-top: 12px;
            font-size: 13px;
            color: #64748b;
            word-break: break-all;
        }
    </style>
</head>
<body>
<div class="container">
    <h1>NCOE Prospective Teachers Digital Readiness Survey (Survey A)</h1>
    <p style="text-align:center; font-size:16px; color:#1e40af;">
        Digital Readiness &amp; Liveware Awareness Survey for Prospective Teachers (Pre-Service)
    </p>

    <!-- ===================== QR CODE SECTION ===================== -->
    <div class="qr-section">
        <h3>📱 Scan this QR Code to open the survey</h3>
        <p>Respondents can scan this code with their phone camera to open the survey link.</p>
        <div id="qrcode"></div>
        <div class="qr-url" id="qrUrlText"></div>
    </div>
    <!-- ========================================================== -->

    <div class="instruction-box">
        <strong>How submission works:</strong><br>
        1. Fill the form and click the button below.<br>
        2. An Excel file (.xlsx) of your answers will be downloaded automatically.<br>
        3. Your email client will open — please <strong>attach the downloaded Excel file</strong> and send it to the researcher at <strong>cletusrodrigo@gmail.com</strong>.
    </div>

    <form id="surveyForm" onsubmit="return false;">
        <!-- PART 1 -->
        <h2>Part 1: Demographics &amp; Access</h2>
        <div class="question"><label>1. Name of the NCOE you are currently enrolled in <span style="color:red">*</span> <span class="error" id="err1">Required</span></label><input type="text" id="q1" placeholder="e.g. Maharagama NCOE"></div>
        <div class="question"><label>2. Your main teaching subject <span style="color:red">*</span> <span class="error" id="err2">Required</span></label><input type="text" id="q2" placeholder="e.g. Mathematics"></div>
        <div class="question"><label>3. Do you own a personal laptop or tablet?</label>
            <select id="q3"><option value="">Select</option><option>Yes — Laptop</option><option>Yes — Tablet</option><option>Both</option><option>No</option></select>
        </div>
        <div class="question"><label>4. How many hours per day do you have access to high-speed internet at your NCOE?</label>
            <select id="q4"><option value="">Select</option><option>Less than 1 hour</option><option>1–2 hours</option><option>3–4 hours</option><option>More than 4 hours</option><option>No access</option></select>
        </div>
        <div class="question"><label>5. Are you aware of the Sri Lanka Education Reforms 2026? (1=Very Unaware → 5=Very Aware)</label><div class="likert"><label><input type="radio" name="q5" value="1">1</label><label><input type="radio" name="q5" value="2">2</label><label><input type="radio" name="q5" value="3">3</label><label><input type="radio" name="q5" value="4">4</label><label><input type="radio" name="q5" value="5">5</label></div></div>
        <div class="question"><label>6. Have you received a personal digital device from the government? (1=No → 5=Yes and use regularly)</label><div class="likert"><label><input type="radio" name="q6" value="1">1</label><label><input type="radio" name="q6" value="2">2</label><label><input type="radio" name="q6" value="3">3</label><label><input type="radio" name="q6" value="4">4</label><label><input type="radio" name="q6" value="5">5</label></div></div>
        <div class="question"><label>7. Do you have access to a Smart Classroom at least once a week? (1=Never → 5=Daily)</label><div class="likert"><label><input type="radio" name="q7" value="1">1</label><label><input type="radio" name="q7" value="2">2</label><label><input type="radio" name="q7" value="3">3</label><label><input type="radio" name="q7" value="4">4</label><label><input type="radio" name="q7" value="5">5</label></div></div>
        <div class="question"><label>8. How would you rate your general computer literacy? (1=Very Low → 5=Expert)</label><div class="likert"><label><input type="radio" name="q8" value="1">1</label><label><input type="radio" name="q8" value="2">2</label><label><input type="radio" name="q8" value="3">3</label><label><input type="radio" name="q8" value="4">4</label><label><input type="radio" name="q8" value="5">5</label></div></div>
        <div class="question"><label>9. Prior to NCOE, did you have formal ICT training? (1=None → 5=Extensive)</label><div class="likert"><label><input type="radio" name="q9" value="1">1</label><label><input type="radio" name="q9" value="2">2</label><label><input type="radio" name="q9" value="3">3</label><label><input type="radio" name="q9" value="4">4</label><label><input type="radio" name="q9" value="5">5</label></div></div>
        <div class="question"><label>10. Is the software provided by the government installed on the devices you use? (1=Not at all → 5=Fully installed and functional)</label><div class="likert"><label><input type="radio" name="q10" value="1">1</label><label><input type="radio" name="q10" value="2">2</label><label><input type="radio" name="q10" value="3">3</label><label><input type="radio" name="q10" value="4">4</label><label><input type="radio" name="q10" value="5">5</label></div></div>

        <!-- PART 2 -->
        <h2>Part 2: Technological Knowledge (TK)</h2>
        <p style="color:#555;font-size:14px;margin:0 0 10px">Rate 1 (Strongly Disagree) to 5 (Strongly Agree)</p>
        <div class="question"><label>11. I am comfortable navigating the operating system (Windows/Linux) provided.</label><div class="likert"><label><input type="radio" name="q11" value="1">1</label><label><input type="radio" name="q11" value="2">2</label><label><input type="radio" name="q11" value="3">3</label><label><input type="radio" name="q11" value="4">4</label><label><input type="radio" name="q11" value="5">5</label></div></div>
        <div class="question"><label>12. I can troubleshoot basic hardware issues (e.g., connecting a projector).</label><div class="likert"><label><input type="radio" name="q12" value="1">1</label><label><input type="radio" name="q12" value="2">2</label><label><input type="radio" name="q12" value="3">3</label><label><input type="radio" name="q12" value="4">4</label><label><input type="radio" name="q12" value="5">5</label></div></div>
        <div class="question"><label>13. I can use office productivity tools (Word, Excel, PowerPoint) efficiently.</label><div class="likert"><label><input type="radio" name="q13" value="1">1</label><label><input type="radio" name="q13" value="2">2</label><label><input type="radio" name="q13" value="3">3</label><label><input type="radio" name="q13" value="4">4</label><label><input type="radio" name="q13" value="5">5</label></div></div>
        <div class="question"><label>14. I am aware of how to use cloud storage (Google Drive / OneDrive) for my assignments.</label><div class="likert"><label><input type="radio" name="q14" value="1">1</label><label><input type="radio" name="q14" value="2">2</label><label><input type="radio" name="q14" value="3">3</label><label><input type="radio" name="q14" value="4">4</label><label><input type="radio" name="q14" value="5">5</label></div></div>
        <div class="question"><label>15. I can install and update educational software independently.</label><div class="likert"><label><input type="radio" name="q15" value="1">1</label><label><input type="radio" name="q15" value="2">2</label><label><input type="radio" name="q15" value="3">3</label><label><input type="radio" name="q15" value="4">4</label><label><input type="radio" name="q15" value="5">5</label></div></div>
        <div class="question"><label>16. I am confident in using the specific LMS provided by the government.</label><div class="likert"><label><input type="radio" name="q16" value="1">1</label><label><input type="radio" name="q16" value="2">2</label><label><input type="radio" name="q16" value="3">3</label><label><input type="radio" name="q16" value="4">4</label><label><input type="radio" name="q16" value="5">5</label></div></div>
        <div class="question"><label>17. I know how to use digital security tools to protect my data.</label><div class="likert"><label><input type="radio" name="q17" value="1">1</label><label><input type="radio" name="q17" value="2">2</label><label><input type="radio" name="q17" value="3">3</label><label><input type="radio" name="q17" value="4">4</label><label><input type="radio" name="q17" value="5">5</label></div></div>
        <div class="question"><label>18. I can manipulate multimedia files (editing images or videos) for my lessons.</label><div class="likert"><label><input type="radio" name="q18" value="1">1</label><label><input type="radio" name="q18" value="2">2</label><label><input type="radio" name="q18" value="3">3</label><label><input type="radio" name="q18" value="4">4</label><label><input type="radio" name="q18" value="5">5</label></div></div>
        <div class="question"><label>19. I am aware of how to use interactive whiteboard software.</label><div class="likert"><label><input type="radio" name="q19" value="1">1</label><label><input type="radio" name="q19" value="2">2</label><label><input type="radio" name="q19" value="3">3</label><label><input type="radio" name="q19" value="4">4</label><label><input type="radio" name="q19" value="5">5</label></div></div>
        <div class="question"><label>20. I find it easy to learn new digital applications.</label><div class="likert"><label><input type="radio" name="q20" value="1">1</label><label><input type="radio" name="q20" value="2">2</label><label><input type="radio" name="q20" value="3">3</label><label><input type="radio" name="q20" value="4">4</label><label><input type="radio" name="q20" value="5">5</label></div></div>

        <!-- PART 3 -->
        <h2>Part 3: Pedagogical Awareness (PK)</h2>
        <p style="color:#555;font-size:14px;margin:0 0 10px">Rate 1 (Strongly Disagree) to 5 (Strongly Agree)</p>
        <div class="question"><label>21. I understand how to manage a classroom where students are using digital devices.</label><div class="likert"><label><input type="radio" name="q21" value="1">1</label><label><input type="radio" name="q21" value="2">2</label><label><input type="radio" name="q21" value="3">3</label><label><input type="radio" name="q21" value="4">4</label><label><input type="radio" name="q21" value="5">5</label></div></div>
        <div class="question"><label>22. I am aware of 'Digital Citizenship' and how to teach it to my students.</label><div class="likert"><label><input type="radio" name="q22" value="1">1</label><label><input type="radio" name="q22" value="2">2</label><label><input type="radio" name="q22" value="3">3</label><label><input type="radio" name="q22" value="4">4</label><label><input type="radio" name="q22" value="5">5</label></div></div>
        <div class="question"><label>23. I can assess student work using digital platforms.</label><div class="likert"><label><input type="radio" name="q23" value="1">1</label><label><input type="radio" name="q23" value="2">2</label><label><input type="radio" name="q23" value="3">3</label><label><input type="radio" name="q23" value="4">4</label><label><input type="radio" name="q23" value="5">5</label></div></div>
        <div class="question"><label>24. I believe digital tools can help me cater to students with different learning speeds.</label><div class="likert"><label><input type="radio" name="q24" value="1">1</label><label><input type="radio" name="q24" value="2">2</label><label><input type="radio" name="q24" value="3">3</label><label><input type="radio" name="q24" value="4">4</label><label><input type="radio" name="q24" value="5">5</label></div></div>
        <div class="question"><label>25. I am aware of the SAMR model (Substitution, Augmentation, Modification, Redefinition) of technology integration.</label><div class="likert"><label><input type="radio" name="q25" value="1">1</label><label><input type="radio" name="q25" value="2">2</label><label><input type="radio" name="q25" value="3">3</label><label><input type="radio" name="q25" value="4">4</label><label><input type="radio" name="q25" value="5">5</label></div></div>
        <div class="question"><label>26. I can design a 40-minute lesson plan that is fully integrated with digital tools.</label><div class="likert"><label><input type="radio" name="q26" value="1">1</label><label><input type="radio" name="q26" value="2">2</label><label><input type="radio" name="q26" value="3">3</label><label><input type="radio" name="q26" value="4">4</label><label><input type="radio" name="q26" value="5">5</label></div></div>
        <div class="question"><label>27. I know how to search for Open Educational Resources (OERs) safely.</label><div class="likert"><label><input type="radio" name="q27" value="1">1</label><label><input type="radio" name="q27" value="2">2</label><label><input type="radio" name="q27" value="3">3</label><label><input type="radio" name="q27" value="4">4</label><label><input type="radio" name="q27" value="5">5</label></div></div>
        <div class="question"><label>28. I am aware of how to facilitate online collaborative projects for students.</label><div class="likert"><label><input type="radio" name="q28" value="1">1</label><label><input type="radio" name="q28" value="2">2</label><label><input type="radio" name="q28" value="3">3</label><label><input type="radio" name="q28" value="4">4</label><label><input type="radio" name="q28" value="5">5</label></div></div>
        <div class="question"><label>29. I understand the ethical implications of using AI in the classroom.</label><div class="likert"><label><input type="radio" name="q29" value="1">1</label><label><input type="radio" name="q29" value="2">2</label><label><input type="radio" name="q29" value="3">3</label><label><input type="radio" name="q29" value="4">4</label><label><input type="radio" name="q29" value="5">5</label></div></div>
        <div class="question"><label>30. I can identify when technology is distracting rather than helping my students.</label><div class="likert"><label><input type="radio" name="q30" value="1">1</label><label><input type="radio" name="q30" value="2">2</label><label><input type="radio" name="q30" value="3">3</label><label><input type="radio" name="q30" value="4">4</label><label><input type="radio" name="q30" value="5">5</label></div></div>

        <!-- PART 4 -->
        <h2>Part 4: Content &amp; Technological Integration (TCK / TPACK)</h2>
        <p style="color:#555;font-size:14px;margin:0 0 10px">Rate 1 (Strongly Disagree) to 5 (Strongly Agree)</p>
        <div class="question"><label>31. I know specific software that can explain complex concepts in my subject area.</label><div class="likert"><label><input type="radio" name="q31" value="1">1</label><label><input type="radio" name="q31" value="2">2</label><label><input type="radio" name="q31" value="3">3</label><label><input type="radio" name="q31" value="4">4</label><label><input type="radio" name="q31" value="5">5</label></div></div>
        <div class="question"><label>32. I can use digital simulations to show practical experiments in my subject.</label><div class="likert"><label><input type="radio" name="q32" value="1">1</label><label><input type="radio" name="q32" value="2">2</label><label><input type="radio" name="q32" value="3">3</label><label><input type="radio" name="q32" value="4">4</label><label><input type="radio" name="q32" value="5">5</label></div></div>
        <div class="question"><label>33. I am aware of how digital tools can make my specific subject more interesting to students.</label><div class="likert"><label><input type="radio" name="q33" value="1">1</label><label><input type="radio" name="q33" value="2">2</label><label><input type="radio" name="q33" value="3">3</label><label><input type="radio" name="q33" value="4">4</label><label><input type="radio" name="q33" value="5">5</label></div></div>
        <div class="question"><label>34. My lecturers at NCOE regularly use digital tools to teach subject-specific content.</label><div class="likert"><label><input type="radio" name="q34" value="1">1</label><label><input type="radio" name="q34" value="2">2</label><label><input type="radio" name="q34" value="3">3</label><label><input type="radio" name="q34" value="4">4</label><label><input type="radio" name="q34" value="5">5</label></div></div>
        <div class="question"><label>35. I feel I can explain my subject better using a smartboard than a chalkboard.</label><div class="likert"><label><input type="radio" name="q35" value="1">1</label><label><input type="radio" name="q35" value="2">2</label><label><input type="radio" name="q35" value="3">3</label><label><input type="radio" name="q35" value="4">4</label><label><input type="radio" name="q35" value="5">5</label></div></div>
        <div class="question"><label>36. I am aware of digital databases relevant to my teaching subject.</label><div class="likert"><label><input type="radio" name="q36" value="1">1</label><label><input type="radio" name="q36" value="2">2</label><label><input type="radio" name="q36" value="3">3</label><label><input type="radio" name="q36" value="4">4</label><label><input type="radio" name="q36" value="5">5</label></div></div>
        <div class="question"><label>37. I can create my own digital content (e.g., quizzes) specific to my subject.</label><div class="likert"><label><input type="radio" name="q37" value="1">1</label><label><input type="radio" name="q37" value="2">2</label><label><input type="radio" name="q37" value="3">3</label><label><input type="radio" name="q37" value="4">4</label><label><input type="radio" name="q37" value="5">5</label></div></div>
        <div class="question"><label>38. I have seen my lecturers demonstrate the software the government is providing.</label><div class="likert"><label><input type="radio" name="q38" value="1">1</label><label><input type="radio" name="q38" value="2">2</label><label><input type="radio" name="q38" value="3">3</label><label><input type="radio" name="q38" value="4">4</label><label><input type="radio" name="q38" value="5">5</label></div></div>
        <div class="question"><label>39. I am aware of how to use data analytics to track student progress in my subject.</label><div class="likert"><label><input type="radio" name="q39" value="1">1</label><label><input type="radio" name="q39" value="2">2</label><label><input type="radio" name="q39" value="3">3</label><label><input type="radio" name="q39" value="4">4</label><label><input type="radio" name="q39" value="5">5</label></div></div>
        <div class="question"><label>40. I feel the provided software is relevant to the national syllabus of my subject.</label><div class="likert"><label><input type="radio" name="q40" value="1">1</label><label><input type="radio" name="q40" value="2">2</label><label><input type="radio" name="q40" value="3">3</label><label><input type="radio" name="q40" value="4">4</label><label><input type="radio" name="q40" value="5">5</label></div></div>

        <!-- PART 5 -->
        <h2>Part 5: Psychological Readiness &amp; Attitudes</h2>
        <p style="color:#555;font-size:14px;margin:0 0 10px">Rate 1 (Strongly Disagree) to 5 (Strongly Agree). Note: Q44 and Q48 are reverse-coded.</p>
        <div class="question"><label>41. I feel excited about the digitalization of Sri Lankan schools.</label><div class="likert"><label><input type="radio" name="q41" value="1">1</label><label><input type="radio" name="q41" value="2">2</label><label><input type="radio" name="q41" value="3">3</label><label><input type="radio" name="q41" value="4">4</label><label><input type="radio" name="q41" value="5">5</label></div></div>
        <div class="question"><label>42. I am worried that technology will make my job as a teacher more difficult.</label><div class="likert"><label><input type="radio" name="q42" value="1">1</label><label><input type="radio" name="q42" value="2">2</label><label><input type="radio" name="q42" value="3">3</label><label><input type="radio" name="q42" value="4">4</label><label><input type="radio" name="q42" value="5">5</label></div></div>
        <div class="question"><label>43. I feel confident leading a Smart Classroom in my future school.</label><div class="likert"><label><input type="radio" name="q43" value="1">1</label><label><input type="radio" name="q43" value="2">2</label><label><input type="radio" name="q43" value="3">3</label><label><input type="radio" name="q43" value="4">4</label><label><input type="radio" name="q43" value="5">5</label></div></div>
        <div class="question"><label>44. I believe the government's investment in hardware is more important than teacher training. <span class="reverse-note">[REVERSE CODED]</span></label><div class="likert"><label><input type="radio" name="q44" value="1">1</label><label><input type="radio" name="q44" value="2">2</label><label><input type="radio" name="q44" value="3">3</label><label><input type="radio" name="q44" value="4">4</label><label><input type="radio" name="q44" value="5">5</label></div></div>
        <div class="question"><label>45. I feel that my lecturers are sufficiently prepared to guide me in digital teaching.</label><div class="likert"><label><input type="radio" name="q45" value="1">1</label><label><input type="radio" name="q45" value="2">2</label><label><input type="radio" name="q45" value="3">3</label><label><input type="radio" name="q45" value="4">4</label><label><input type="radio" name="q45" value="5">5</label></div></div>
        <div class="question"><label>46. I am willing to spend extra time learning new software for my profession.</label><div class="likert"><label><input type="radio" name="q46" value="1">1</label><label><input type="radio" name="q46" value="2">2</label><label><input type="radio" name="q46" value="3">3</label><label><input type="radio" name="q46" value="4">4</label><label><input type="radio" name="q46" value="5">5</label></div></div>
        <div class="question"><label>47. I believe that digital tools will improve the quality of education in rural schools.</label><div class="likert"><label><input type="radio" name="q47" value="1">1</label><label><input type="radio" name="q47" value="2">2</label><label><input type="radio" name="q47" value="3">3</label><label><input type="radio" name="q47" value="4">4</label><label><input type="radio" name="q47" value="5">5</label></div></div>
        <div class="question"><label>48. I fear that technology might eventually replace the need for traditional teaching. <span class="reverse-note">[REVERSE CODED]</span></label><div class="likert"><label><input type="radio" name="q48" value="1">1</label><label><input type="radio" name="q48" value="2">2</label><label><input type="radio" name="q48" value="3">3</label><label><input type="radio" name="q48" value="4">4</label><label><input type="radio" name="q48" value="5">5</label></div></div>
        <div class="question"><label>49. I feel supported by my NCOE administration in my digital learning journey.</label><div class="likert"><label><input type="radio" name="q49" value="1">1</label><label><input type="radio" name="q49" value="2">2</label><label><input type="radio" name="q49" value="3">3</label><label><input type="radio" name="q49" value="4">4</label><label><input type="radio" name="q49" value="5">5</label></div></div>
        <div class="question"><label>50. I am ready to be a Digital Leader in the school I am assigned to in 2026.</label><div class="likert"><label><input type="radio" name="q50" value="1">1</label><label><input type="radio" name="q50" value="2">2</label><label><input type="radio" name="q50" value="3">3</label><label><input type="radio" name="q50" value="4">4</label><label><input type="radio" name="q50" value="5">5</label></div></div>

        <div class="question"><label>51. In your opinion, what single institutional or policy change would most improve your digital readiness as a future teacher in Sri Lanka? Please explain briefly.</label><textarea id="q51" rows="4" placeholder="Your response..."></textarea></div>

        <div class="btn-row">
            <button class="btn-submit" type="button" onclick="submitSurvey()">
                📊 Generate Excel &amp; Email to Researcher
            </button>
        </div>
        <p class="note">Your response will be prepared as an Excel file and sent to <strong>cletusrodrigo@gmail.com</strong></p>
        <div class="success-banner" id="successBanner"></div>
    </form>
</div>

<script>
/* ============================================================
   IMPORTANT: Replace the URL below with the real public link
   where you host this HTML file (GitHub Pages, Netlify, etc.)
   ============================================================ */
const SURVEY_URL = "https://NCOE.github.io/ncoe-survey/index.html";   // ← CHANGE THIS

// Generate QR Code when page loads
window.addEventListener("DOMContentLoaded", function () {
    const qrDiv = document.getElementById("qrcode");
    const urlText = document.getElementById("qrUrlText");

    // Clear any previous content
    qrDiv.innerHTML = "";

    // Create QR code (size 180 × 180)
    new QRCode(qrDiv, {
        text: SURVEY_URL,
        width: 180,
        height: 180,
        colorDark: "#1e3a8a",
        colorLight: "#ffffff",
        correctLevel: QRCode.CorrectLevel.H
    });

    urlText.textContent = "Link: " + SURVEY_URL;
});

/* ===================== SURVEY LOGIC ===================== */
const HEADERS = [
    "Timestamp","Q1_NCOE","Q2_Subject","Q3_DeviceOwnership","Q4_InternetHours",
    "Q5_ReformsAware","Q6_GovDevice","Q7_SmartClass","Q8_CompLiteracy","Q9_PriorICT","Q10_GovSoftware",
    "Q11_TK_OS","Q12_TK_Hardware","Q13_TK_Office","Q14_TK_Cloud","Q15_TK_Install","Q16_TK_LMS",
    "Q17_TK_Security","Q18_TK_Multimedia","Q19_TK_IWB","Q20_TK_LearnNew",
    "Q21_PK_ClassMgmt","Q22_PK_DigCitizenship","Q23_PK_Assessment","Q24_PK_DiffSpeed","Q25_PK_SAMR",
    "Q26_PK_LessonPlan","Q27_PK_OER","Q28_PK_Collab","Q29_PK_AIEthics","Q30_PK_Distraction",
    "Q31_TCK_Software","Q32_TCK_Simulations","Q33_TCK_Interest","Q34_TCK_LecturersUse","Q35_TCK_Smartboard",
    "Q36_TCK_Databases","Q37_TCK_CreateContent","Q38_TCK_Demo","Q39_TCK_Analytics","Q40_TCK_Relevance",
    "Q41_PR_Excited","Q42_PR_Worried","Q43_PR_Confident","Q44_PR_HardwareOverTraining","Q45_PR_LecturersReady",
    "Q46_PR_WillingLearn","Q47_PR_RuralImprove","Q48_PR_ReplaceFear","Q49_PR_AdminSupport","Q50_PR_DigitalLeader",
    "Q51_OpenEnded"
];

function getVal(id) {
    const el = document.getElementById(id);
    return el ? el.value.trim() : "";
}
function getRadio(name) {
    const sel = document.querySelector(`input[name="${name}"]:checked`);
    return sel ? sel.value : "";
}

function collectValues() {
    return [
        new Date().toLocaleString("en-LK"),
        getVal("q1"), getVal("q2"), getVal("q3"), getVal("q4"),
        getRadio("q5"), getRadio("q6"), getRadio("q7"), getRadio("q8"), getRadio("q9"), getRadio("q10"),
        getRadio("q11"), getRadio("q12"), getRadio("q13"), getRadio("q14"), getRadio("q15"),
        getRadio("q16"), getRadio("q17"), getRadio("q18"), getRadio("q19"), getRadio("q20"),
        getRadio("q21"), getRadio("q22"), getRadio("q23"), getRadio("q24"), getRadio("q25"),
        getRadio("q26"), getRadio("q27"), getRadio("q28"), getRadio("q29"), getRadio("q30"),
        getRadio("q31"), getRadio("q32"), getRadio("q33"), getRadio("q34"), getRadio("q35"),
        getRadio("q36"), getRadio("q37"), getRadio("q38"), getRadio("q39"), getRadio("q40"),
        getRadio("q41"), getRadio("q42"), getRadio("q43"), getRadio("q44"), getRadio("q45"),
        getRadio("q46"), getRadio("q47"), getRadio("q48"), getRadio("q49"), getRadio("q50"),
        getVal("q51")
    ];
}

function validate() {
    const n = getVal("q1"), s = getVal("q2");
    document.getElementById("err1").style.display = n ? "none" : "inline";
    document.getElementById("err2").style.display = s ? "none" : "inline";
    if (!n || !s) {
        alert("Please fill in Question 1 (NCOE) and Question 2 (Subject) before submitting.");
        return false;
    }
    return true;
}

function showBanner(msg) {
    const b = document.getElementById("successBanner");
    b.innerHTML = msg;
    b.style.display = "block";
    b.scrollIntoView({ behavior: "smooth" });
}

function submitSurvey() {
    if (!validate()) return;

    if (typeof XLSX === "undefined") {
        alert("Excel library failed to load. Please check your internet connection and try again.");
        return;
    }

    const values = collectValues();
    const ncoe = getVal("q1").replace(/[^a-zA-Z0-9]/g, "_").substring(0, 30) || "NCOE";
    const dateStr = new Date().toISOString().slice(0, 10);
    const fileName = `Prospective_Teachers_Survey_${ncoe}_${dateStr}.xlsx`;

    // Create Excel workbook
    const wsData = [HEADERS, values];
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    ws["!cols"] = HEADERS.map(h => ({ wch: Math.max(12, Math.min(40, h.length + 2)) }));

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Survey Response");

    // Download Excel file
    XLSX.writeFile(wb, fileName);

    // Open email client
    const bodyText =
`Dear Researcher,

Please find attached the Excel file containing my response to the
NCOE Prospective Teachers Digital Readiness Survey (Survey A).

File name: ${fileName}

Submitted on: ${new Date().toLocaleString("en-LK")}

Thank you.`;

    const mailtoURL =
        `mailto:cletusrodrigo@gmail.com` +
        `?subject=${encodeURIComponent("Prospective_Teachers_Survey_" + ncoe + "_" + dateStr)}` +
        `&body=${encodeURIComponent(bodyText)}`;

    window.location.href = mailtoURL;

    showBanner(
        `✅ <strong>Excel file downloaded!</strong><br>
         File: <strong>${fileName}</strong><br><br>
         Your email client has opened.<br>
         Please <strong>attach the downloaded Excel file</strong> and send it to<br>
         <strong>cletusrodrigo@gmail.com</strong>`
    );
}
</script>
</body>
</html>
