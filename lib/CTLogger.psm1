using module .\CTTimer.psm1;
Class CTLogger {
    [CTTimer] $cttimer;
    [System.Collections.Generic.Dictionary[string,psobject]] $tasks;
    [string] $fileName;
    [string] $connectionType;
    [string] $task;


    CTLogger($connectionType,$task,$fileName) {
        $this.cttimer = New-Object CTTimer;
        $this.tasks = New-Object 'System.Collections.Generic.Dictionary[string,psobject]';
        $this.fileName = $fileName;
        $this.connectionType = $connectionType
        $this.task = $task;
        $this.prepareLogFile()
    }

	[void] startSubtask($subtaskName){
        $this.startSubtask($subtaskName, '');
    }

	<# $subtaskName [string] name of subtask
		$additionalInfo [string] additional information
	#>
    [void] startSubtask($subtaskName, $additionalInfo){
        if($this.tasks.ContainsKey($subtaskName) -eq $false) {
            $this.tasks.Add($subtaskName,
                [psobject]@{
				   timer = New-Object CTTimer
				   additionalInfo = [string]$additionalInfo
                });
        } else {
            Throw [System.Exception]("subtask already started:" + $subtaskName);
        }
    }

    [void] endSubtask($subtaskName,$status,$message) {
        if($this.tasks.ContainsKey($subtaskName) -eq $true) {
            $subtask = $this.tasks[$subtaskName];
            $subtask.timer.stop();
            $subtask.status = $status;
            $subtask.message = $message;
            $subtask.name = $subtaskName;
            $this.logToFile($subtask);
            $this.tasks.Remove($subtaskName);
        } else {
            Throw [System.Exception]("subtask don't exists:" + $subtaskName);
        }
    }

    [void] logToFile($subtask){
        $content = [string]::Format("{0}; {1}; {2}; {3}; {4}; {5}; {6}; {7}, {8}",
        $this.connectionType, $this.task,$subtask.name, $subtask.timer.startDateTime, $subtask.timer.endDateTime,  $subtask.timer.totalSeconds(), $subtask.status, $subtask.additionalInfo, $subtask.message);
        Add-Content -path $this.fileName $content;
    }

    [void] prepareLogFile(){
        if(Test-Path -Path $this.fileName){
            return;
        }

        $content = [string]::Format("Connection Type; Task; Subtask; Start Time; End Time; Duration [s]; Task Status; Remarks; Addittional Info");
        Add-Content -Path $this.fileName $content;
        
    }

    

}