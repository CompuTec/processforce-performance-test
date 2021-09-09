


Class CTTimer {
    [DateTime] $startDateTime;
    [DateTime] $endDateTime;
    [DateTime] $prevStepDateTime;
    [DateTime] $currentDateTime;
    [bool] $started;
    CTTimer() {
        $this.startDateTime = Get-Date;
        $this.prevStepDateTime = $this.startDateTime;
        $this.started = $true;
    }

    [double] round(){
        if($this.started -eq $false){
            $this.restart();
        }
        $this.currentDateTime = Get-Date;
        $seconds = ($this.currentDateTime-$this.prevStepDateTime).TotalSeconds;
        $this.prevStepDateTime = $this.currentDateTime;
        return $seconds;
    }

    [double] totalSeconds(){
        if($this.started -eq $false){
            $end = $this.endDateTime;
        } else {
            $end = Get-Date;
        }
        $seconds = ($end - $this.startDateTime).TotalSeconds;
        return $seconds;
    }

    [void] stop(){
        $this.endDateTime =  Get-Date;
        $this.started = $false;
    }

    [void] restart() {
        $this.startDateTime = Get-Date;
        $this.prevStepDateTime = $this.startDateTime;
        $this.started = $true;
    }



}
    