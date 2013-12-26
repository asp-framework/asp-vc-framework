<% Response.Buffer = True %>
<!DOCTYPE html>
<html>
<head>
    <title><%= title %></title>
    <meta charset="utf-8">
    <meta name="description" content="SE For ASP" />
    <meta name="keywords" content="SE, ASP, VBScript" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <link rel="shortcut icon" href="favicon.ico" />
    <link rel="icon" href="favicon.ico" />
    <link rel="bookmark" href="favicon.ico" />
    <link rel="stylesheet" href="http://cdn.bootcss.com/twitter-bootstrap/3.0.3/css/bootstrap.min.css" />
    <% '<!-- #contentStartToDo -->' %>
</head>
<% Response.Flush() %>
<body>

    <!-- 导航条 -->
    <nav class="navbar navbar-inverse" role="navigation">
        <div class="container">
            <div class="navbar-header">
                <a class="navbar-brand" href="./">SE&nbsp;For&nbsp;ASP</a>
            </div>
            <div class="collapse navbar-collapse">

            </div>
        </div>
    </nav>
    <!-- /导航条 -->

    <% '<!-- #content -->' %>

    <% Response.Flush() %>
    <!-- Js -->
    <script src="http://cdn.bootcss.com/jquery/1.10.2/jquery.min.js"></script>
    <script src="http://cdn.bootcss.com/twitter-bootstrap/3.0.3/js/bootstrap.min.js"></script>
    <% '<!-- #contentEndToDo -->' %>
    <!-- /Js -->

</body>
</html>