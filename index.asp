<%@ Language="VBScript" %>
<% Option Explicit %>
<%
'=============================================================
' Classic ASP Calculator (server-side safety net)
' This block provides optional validation and calculation if
' a POST request is made. The UI works client-side in real-time.
'=============================================================
Dim serverResult, serverError
serverResult = ""
serverError = ""

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim rawExpression, sanitizedExpression
    rawExpression = Trim(Request.Form("expression"))

    If rawExpression = "" Then
        serverError = "Please enter a calculation."
    Else
        sanitizedExpression = Replace(rawExpression, " ", "")

        ' Allow only digits, decimal points, and operators.
        Dim i, ch
        For i = 1 To Len(sanitizedExpression)
            ch = Mid(sanitizedExpression, i, 1)
            If InStr("0123456789.+-*/", ch) = 0 Then
                serverError = "Invalid characters detected."
                Exit For
            End If
        Next

        If serverError = "" Then
            ' Prevent division by zero (simple check).
            If InStr(sanitizedExpression, "/0") > 0 Then
                serverError = "Division by zero is not allowed."
            Else
                On Error Resume Next
                serverResult = CStr(EvaluateExpression(sanitizedExpression))
                If Err.Number <> 0 Then
                    serverError = "Unable to evaluate the expression."
                    Err.Clear
                End If
                On Error GoTo 0
            End If
        End If
    End If
End If

' Safe, minimal expression evaluation for basic arithmetic.
Function EvaluateExpression(expression)
    ' NOTE: VBScript Eval is not available in Classic ASP.
    ' This parser handles simple left-to-right evaluation.
    Dim numbers(), operators()
    Dim numberCount, operatorCount
    Dim i, current, buffer

    ReDim numbers(0)
    ReDim operators(0)
    numberCount = 0
    operatorCount = 0
    buffer = ""

    For i = 1 To Len(expression)
        current = Mid(expression, i, 1)
        If InStr("0123456789.", current) > 0 Then
            buffer = buffer & current
        Else
            If buffer <> "" Then
                numberCount = numberCount + 1
                ReDim Preserve numbers(numberCount)
                numbers(numberCount) = CDbl(buffer)
                buffer = ""
            End If
            operatorCount = operatorCount + 1
            ReDim Preserve operators(operatorCount)
            operators(operatorCount) = current
        End If
    Next

    If buffer <> "" Then
        numberCount = numberCount + 1
        ReDim Preserve numbers(numberCount)
        numbers(numberCount) = CDbl(buffer)
    End If

    Dim resultValue, idx
    resultValue = numbers(1)
    For idx = 1 To operatorCount
        Select Case operators(idx)
            Case "+"
                resultValue = resultValue + numbers(idx + 1)
            Case "-"
                resultValue = resultValue - numbers(idx + 1)
            Case "*"
                resultValue = resultValue * numbers(idx + 1)
            Case "/"
                If numbers(idx + 1) = 0 Then
                    resultValue = 0
                Else
                    resultValue = resultValue / numbers(idx + 1)
                End If
        End Select
    Next

    EvaluateExpression = resultValue
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Classic ASP Calculator</title>
    <style>
        :root {
            color-scheme: light dark;
            --bg: #f4f6fb;
            --card: #ffffff;
            --text: #1f2933;
            --muted: #6b7c93;
            --accent: #4c6ef5;
            --accent-dark: #364fc7;
            --danger: #e03131;
        }

        * {
            box-sizing: border-box;
        }

        body {
            margin: 0;
            font-family: "Segoe UI", system-ui, -apple-system, sans-serif;
            background: var(--bg);
            color: var(--text);
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 32px;
        }

        main {
            width: min(420px, 100%);
            background: var(--card);
            border-radius: 20px;
            padding: 24px;
            box-shadow: 0 20px 40px rgba(31, 41, 51, 0.15);
        }

        header {
            display: flex;
            flex-direction: column;
            gap: 6px;
            margin-bottom: 20px;
        }

        header h1 {
            font-size: 1.5rem;
            margin: 0;
        }

        header p {
            margin: 0;
            color: var(--muted);
            font-size: 0.95rem;
        }

        .display {
            background: #0f172a;
            color: #f8fafc;
            border-radius: 12px;
            padding: 16px;
            display: grid;
            gap: 6px;
            margin-bottom: 18px;
        }

        .display .expression {
            font-size: 0.95rem;
            min-height: 22px;
            color: #cbd5f5;
        }

        .display .result {
            font-size: 2rem;
            font-weight: 600;
            text-align: right;
        }

        .keypad {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 12px;
        }

        button {
            border: none;
            border-radius: 12px;
            padding: 14px 10px;
            font-size: 1.05rem;
            font-weight: 600;
            background: #e2e8f0;
            color: #111827;
            cursor: pointer;
            transition: transform 0.1s ease, background 0.2s ease;
        }

        button:active {
            transform: scale(0.98);
        }

        button.operator {
            background: #dbe4ff;
            color: var(--accent-dark);
        }

        button.equal {
            background: var(--accent);
            color: #fff;
            grid-column: span 2;
        }

        button.clear {
            background: #ffe3e3;
            color: var(--danger);
        }

        .status {
            margin-top: 16px;
            padding: 10px 12px;
            border-radius: 10px;
            font-size: 0.9rem;
            display: none;
        }

        .status.error {
            background: #ffe3e3;
            color: #c92a2a;
        }

        .status.success {
            background: #d3f9d8;
            color: #2b8a3e;
        }

        @media (max-width: 480px) {
            main {
                padding: 18px;
            }

            .display .result {
                font-size: 1.6rem;
            }
        }
    </style>
</head>
<body>
    <main>
        <header>
            <h1>Classic ASP Calculator</h1>
            <p>Fast, responsive arithmetic with built-in validation.</p>
        </header>

        <section class="display" aria-live="polite">
            <div class="expression" id="expression">0</div>
            <div class="result" id="result">0</div>
        </section>

        <section class="keypad" aria-label="Calculator keypad">
            <button type="button" class="clear" data-action="clear">AC</button>
            <button type="button" class="operator" data-action="delete">DEL</button>
            <button type="button" class="operator" data-action="percent">%</button>
            <button type="button" class="operator" data-value="/">÷</button>

            <button type="button" data-value="7">7</button>
            <button type="button" data-value="8">8</button>
            <button type="button" data-value="9">9</button>
            <button type="button" class="operator" data-value="*">×</button>

            <button type="button" data-value="4">4</button>
            <button type="button" data-value="5">5</button>
            <button type="button" data-value="6">6</button>
            <button type="button" class="operator" data-value="-">−</button>

            <button type="button" data-value="1">1</button>
            <button type="button" data-value="2">2</button>
            <button type="button" data-value="3">3</button>
            <button type="button" class="operator" data-value="+">+</button>

            <button type="button" data-action="toggle-sign">±</button>
            <button type="button" data-value="0">0</button>
            <button type="button" data-value=".">.</button>
            <button type="button" class="equal" data-action="equals">=</button>
        </section>

        <div class="status" id="status"></div>

        <form id="server-form" method="post" style="display:none">
            <input type="hidden" name="expression" id="server-expression" value="" />
        </form>
    </main>

    <script>
        (function () {
            "use strict";

            var expressionEl = document.getElementById("expression");
            var resultEl = document.getElementById("result");
            var statusEl = document.getElementById("status");
            var serverExpressionEl = document.getElementById("server-expression");

            var expression = "";

            function updateDisplay(result, isError) {
                expressionEl.textContent = expression || "0";
                resultEl.textContent = result;
                if (isError) {
                    statusEl.textContent = result;
                    statusEl.className = "status error";
                    statusEl.style.display = "block";
                } else {
                    statusEl.textContent = "";
                    statusEl.className = "status";
                    statusEl.style.display = "none";
                }
            }

            function isOperator(char) {
                return ["+", "-", "*", "/"].indexOf(char) !== -1;
            }

            function sanitizeExpression(value) {
                return value.replace(/[^0-9+\-*/.]/g, "");
            }

            function hasInvalidDecimal(segment) {
                return (segment.match(/\./g) || []).length > 1;
            }

            function validateExpression(value) {
                if (!value) {
                    return "Please enter a calculation.";
                }

                if (/[^0-9+\-*/.]/.test(value)) {
                    return "Invalid characters detected.";
                }

                if (/\/+0(?![0-9.])/.test(value)) {
                    return "Division by zero is not allowed.";
                }

                var parts = value.split(/[+\-*/]/);
                for (var i = 0; i < parts.length; i += 1) {
                    if (hasInvalidDecimal(parts[i])) {
                        return "Invalid decimal format.";
                    }
                }

                return "";
            }

            function calculate(value) {
                var sanitized = sanitizeExpression(value);
                var validationError = validateExpression(sanitized);
                if (validationError) {
                    updateDisplay(validationError, true);
                    return;
                }

                try {
                    var computed = Function("'use strict'; return (" + sanitized + ")")();
                    if (!isFinite(computed)) {
                        updateDisplay("Result is undefined.", true);
                        return;
                    }

                    updateDisplay(formatResult(computed), false);
                } catch (error) {
                    updateDisplay("Unable to evaluate the expression.", true);
                }
            }

            function formatResult(value) {
                var rounded = Math.round((value + Number.EPSILON) * 1000000000) / 1000000000;
                return rounded.toString();
            }

            function appendValue(value) {
                if (expression === "" && value === "0") {
                    updateDisplay("0", false);
                    return;
                }

                var lastChar = expression.slice(-1);
                if (isOperator(lastChar) && isOperator(value)) {
                    expression = expression.slice(0, -1) + value;
                } else {
                    expression += value;
                }

                calculate(expression);
            }

            function handleAction(action) {
                switch (action) {
                    case "clear":
                        expression = "";
                        updateDisplay("0", false);
                        break;
                    case "delete":
                        expression = expression.slice(0, -1);
                        calculate(expression);
                        break;
                    case "toggle-sign":
                        if (!expression) {
                            expression = "-";
                        } else {
                            expression = "-" + expression;
                        }
                        calculate(expression);
                        break;
                    case "percent":
                        if (expression) {
                            expression = expression + "/100";
                            calculate(expression);
                        }
                        break;
                    case "equals":
                        calculate(expression);
                        break;
                    default:
                        break;
                }
            }

            document.querySelectorAll("button").forEach(function (button) {
                button.addEventListener("click", function () {
                    var value = button.getAttribute("data-value");
                    var action = button.getAttribute("data-action");
                    if (value) {
                        appendValue(value);
                    } else if (action) {
                        handleAction(action);
                    }
                });
            });

            document.addEventListener("keydown", function (event) {
                var key = event.key;
                if (/^[0-9]$/.test(key)) {
                    appendValue(key);
                } else if (["+", "-", "*", "/", "."].indexOf(key) !== -1) {
                    appendValue(key);
                } else if (key === "Enter") {
                    event.preventDefault();
                    handleAction("equals");
                } else if (key === "Backspace") {
                    handleAction("delete");
                } else if (key === "Escape") {
                    handleAction("clear");
                }
            });

            // Optional server-side evaluation hook for future use.
            if (serverExpressionEl) {
                serverExpressionEl.value = expression;
            }

            <% If serverError <> "" Then %>
                expression = "<%=Replace(Replace(rawExpression, "\"", "\\\""), "'", "\\'")%>";
                updateDisplay("<%=serverError%>", true);
            <% ElseIf serverResult <> "" Then %>
                expression = "<%=Replace(Replace(rawExpression, "\"", "\\\""), "'", "\\'")%>";
                updateDisplay("<%=serverResult%>", false);
            <% End If %>
        })();
    </script>
</body>
</html>
