<!DOCTYPE html>
<html lang="zh-Hant">

<head>
    <meta charset="utf-8">
    <title>宏謙科技實業有限公司 - 入案管理系統</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">

    <!-- Bootstrap -->
    <link rel="stylesheet" href="./css/bootstrap-5.3.0.min.css">

    <style>
        html,
        body {
            height: 100%;
            margin: 0;
            font-family: 'Segoe UI', sans-serif;
        }

        body {
            background: url("./Image/bg.png") no-repeat center center fixed;
            background-size: cover;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
        }

        .overlay {
            position: fixed;
            inset: 0;
            background: rgba(0, 0, 0, 0.003);
            z-index: -1;
        }

        .main-container {
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            min-height: 100vh;
            position: relative;
            z-index: 1;
        }

        .system-title {
            text-align: center;
            margin-bottom: 25px;
        }

        .main-title {
            margin: 0;
            font-size: 1.7rem;
            font-weight: 700;
            color: #1e3c72;
            letter-spacing: 4px;
        }

        .second-title {
            margin-top: 10px;
            font-size: 1.5rem;
            font-weight: 500;
            color: #4a90e2;
            letter-spacing: 3px;
        }
        
        .login-box {
            width: 100%;
            max-width: 380px;
            padding: 30px;
            border-radius: 12px;
            box-shadow: 0 4px 15px rgba(0, 123, 255, 0.25);
            background-color: white;
        }
        
        .login-box h2 {
            text-align: center;
            margin-bottom: 20px;
            color: #1e88e5;
        }

        .form-control:focus {
            border-color: #1e88e5;
            box-shadow: 0 0 5px rgba(30, 136, 229, 0.5);
        }

        button {
            width: 100%;
            background-color: #1e88e5;
            border: none;
        }

        button:hover {
            background-color: #1565c0;
        }

        .info-box {
            margin-top: 20px;
            font-size: 0.9rem;
            color: #333;
            border: 1px solid #1e88e5;
            border-radius: 8px;
            padding: 15px 20px;
            background-color: #e3f2fd;
            line-height: 1.6;
        }

        .info-box ul {
            padding-left: 1.2em;
        }

        .copyright {
            position: fixed;
            bottom: 8px;
            left: 50%;
            transform: translateX(-50%);
            font-size: 0.85rem;
            color: #f0f0f0;
            text-shadow: 0 1px 4px rgba(0, 0, 0, 0.8);
            user-select: none;
        }
    </style>
</head>

<body>
    <div class="overlay"></div>

    <div class="main-container">
        <div class="login-box">
            <div class="system-title">
            <div class="main-title">宏謙科技實業有限公司</div>
            <div class="second-title">入案管理系統</div>
            </div>

            <form name="myForm" method="post" action="UserLogin_Contral.asp" onsubmit="return User_Login();">

            <div class="form-group">
                <label class="form-label">使用者帳號</label>
                <input name="MemberID"
                       type="text"
                       class="form-control"
                       maxlength="10"
                       onkeyup="value=value.toUpperCase()"
                       required>
            </div>

            <div class="form-group">
                <label class="form-label">密碼</label>
                <input name="MemberPW"
                       type="password"
                       class="form-control"
                       required>
            </div>

            <button type="submit" class="btn btn-primary w-100 mt-4">
                登入
            </button>
            </form>

            <div class="info-box">
                <strong>注意事項：</strong>
                <ul>
                    <li>密碼須包含至少8碼以上</li>
                    <li>需同時具備四種字元(英文大小寫、數字、特殊符號)中的三種</li>
                </ul>
            </div>
        </div>
    </div>
    <div class="copyright">COPYRIGHT © 2025</div>

</body>
</html>
