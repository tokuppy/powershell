　function SqlExequteNonQuery($SqlCommand,$Server,$Database,$CommnadType){
    try{

        $conStr = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
        $conStr["Data Source"] = $Server
        $conStr["Initial Catalog"] = $Database
        $conStr["User ID"]="*id*"
        $conStr["Password"]="*pass*"
        $conStr["Connect Timeout"] = 300

        $con = New-Object System.Data.SqlClient.SqlConnection
        $con.ConnectionString = $conStr
        

        $cmd =New-Object System.Data.SqlClient.SqlCommand
        $cmd.CommandTimeout = 300
        $cmd.CommandType = [System.Data.CommandType]::$CommnadType
        $cmd.Connection = $con
        $cmd.CommandText = $SqlCommand

        $con.Open()
        $result = $cmd.ExecuteNonQuery()

        $con.Close()
        $con.Dispose()
        $cmd.Dispose()

        return $result
                
    }

    catch [Exception]{

        Write-Host �G���[���e�F$_
        Pause

    }

    finally{

        $con.Close()
        $con.Dispose()
        $cmd.Dispose()

    }

}

function SqlExequteScalar($SqlCommand,$Server,$Database,$CommnadType){
    try{

        $conStr = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
        $conStr["Data Source"] = $Server
        $conStr["Initial Catalog"] = $Database
        $conStr["User ID"]="*id*"
        $conStr["Password"]="*pass*"
        $conStr["Connect Timeout"] = 300

        $con = New-Object System.Data.SqlClient.SqlConnection
        $con.ConnectionString = $conStr
       
        $cmd =New-Object System.Data.SqlClient.SqlCommand
        $cmd.CommandTimeout = 300
        $cmd.CommandType = [System.Data.CommandType]::$CommnadType
        $cmd.Connection = $con
        $cmd.CommandText = $SqlCommand

        $con.Open()
        $result = $cmd.ExecuteScalar()

        $con.Close()
        $con.Dispose()
        $cmd.Dispose()

        return $result
                
    }

    catch [Exception]{

        Write-Host �G���[���e�F$_
        Pause

    }

    finally{

        $con.Close()
        $con.Dispose()
        $cmd.Dispose()

    }

}


function SqlExequteGetData($SqlCommand,$Server,$Database,$CommnadType){
    try{

        $conStr = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
        $conStr["Data Source"] = $Server
        $conStr["Initial Catalog"] = $Database
        $conStr["User ID"]="*id*"
        $conStr["Password"]="*pass*"
        $conStr["Connect Timeout"] = 300

        $con = New-Object System.Data.SqlClient.SqlConnection
        $con.ConnectionString = $conStr

        $cmd =New-Object System.Data.SqlClient.SqlCommand
        $cmd.CommandTimeout = 300
        $cmd.CommandType = [System.Data.CommandType]::$CommnadType
        $cmd.Connection = $con
        $cmd.CommandText = $SqlCommand

        $con.Open()

        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
        $dataset = New-Object System.Data.DataSet
        $adapter.Fill($dataset)
        $result = $dataset

        return $result
                
    }

    catch [Exception]{

        Write-Host �G���[���e�F$_
        Pause

    }

    finally{

        $con.Close()
        $con.Dispose()
        $cmd.Dispose()

    }

}



try{


$result1 = SqlExequteNonQuery 'TestStored'  '(local)\SQLEXPRESS'  'test' StoredProcedure

$result2 = SqlExequteGetData 'TestStoredReader'  '(local)\SQLEXPRESS'  'test' StoredProcedure
$result.Tables[0] | Out-GridView

$result3 =  SqlExequteScalar 'TestStoredScalar'  '(local)\SQLEXPRESS'  'test' StoredProcedure
Write-Host $result2


}
catch [Exception]{

    Write-Host '�G���[���������܂���'
    Write-Host "�G���[���e�F$_"
    Pause
}
