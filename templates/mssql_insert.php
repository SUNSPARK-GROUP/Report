<html>
    <head>
        <title></title>
        <meta http-equiv="content-type" content="text/html" charset="UTF-8"/>
    </head>
    <body>
    <?php
        $serverName="192.168.0.214";
        $connectionInfo=array("Database"=>"erps", "UID"=>"apuser", "PWD"=>"0920799339", "CharacterSet"=>"UTF-8");
        $conn=sqlsrv_connect($serverName, $connectionInfo);
      
        $userID=$_POST['userID'];
        $userName=$_POST['userName'];
      	$userName=$_POST['password'];
        echo "輸入值1：".$userID;
        echo "<br />";
        echo "輸入值2：".$userName;
        echo "<br /><br />";
      
       $sql="INSERT INTO good_idea(userID,userName) VALUES('$userID','$userName')";
       $query=sqlsrv_query($conn,$sql)or die("sql error".sqlsrv_errors());
     
       $sql2="select * from good_idea";
       $result=sqlsrv_query($conn,$sql2)or die("sql error".sqlsrv_errors());
      
       echo "讀取good_idea的值：<br />";
       while($row=sqlsrv_fetch_array($result)){
             echo ("<table border=1px><tr>");
             echo ("<td>員工編號：").$row["userID"].("</td>");
             echo ("<td>員工姓名：").$row["userName"].("</td>");
             echo ("</tr></table>");
             echo ("<hr />");
       }
    ?>
    </body>
</html>