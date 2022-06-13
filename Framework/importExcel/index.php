<!DOCTYPE html>
<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
        <title></title>
    </head>
    <body>
        <?php
        require_once './ObjectLineExcel.php';
        $PahtExcel = "C:/Users/RemOpt/Desktop/tempo.xls";
        $objectExcel = new ObjectLineExcel(4, $PahtExcel);
        $objectGet = $objectExcel->getMemory();

        echo "<table border='1'>";
        foreach ($objectGet as $lines) {
            echo "<tr>";
            foreach ($lines as $colun) {
                 echo "<td>" . $colun. "</td>";
            }
            echo "</tr>";
        }
        echo "</table>";
       
        ?>
    </body>
</html>
