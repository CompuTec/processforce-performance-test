


Class CTProgress {
    [int] $totalRows;
    [int] $iteration
    [int] $progress;


    CTProgress($total) {
        if ($total -gt 1) {
            $this.totalRows = $total
        }
        else {
            $this.totalRows = 1
        }
        $this.progress = 0;
        $this.iteration = 0;
    }

    [void] next(){
        $this.iteration++;
        $this.checkProgress();
    }

    [void] checkProgress(){
        $currentProgress = [math]::Round(($this.iteration * 100) / $this.totalRows);
        if($currentProgress -gt $this.progress){
            $this.progress = $currentProgress;
            if(($this.progress%2) -eq 0){
                Write-Host ([string]::Format("{0}% ",$this.progress)) -NoNewline
            }
        }
    }

}