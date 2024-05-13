if(window.location.href.includes('https://hub.exent.com.br/projects//leads')){
        var intervalo = setInterval(()=>{
            try{
                document.querySelector('[aria-controls="lead-table"]').click();
                clearInterval(intervalo)
            }catch{
            }
        })
        setTimeout(()=>{
            location.reload();
        },300000)
}else if(window.location.href.includes('drive.google.com/drive/folders/?fazerdownloads')){
    
        document.querySelectorAll('div[data-tooltip="Download"]').forEach(e=>{
            var intervalo = setInterval(()=>{
                try{
                    e.click()
                    clearInterval(intervalo)
                } catch(error){
                    console.log(error)
                }
            })
        })

    setTimeout(()=>{
        location.reload();
    },150000)
}if(window.location.href.includes('drive.google.com/drive/folders/?fazerdownloads')){
    
    document.querySelectorAll('div[data-tooltip="Download"]').forEach(e=>{
        var intervalo = setInterval(()=>{
            try{
                e.click()
                clearInterval(intervalo)
            } catch(error){
                console.log(error)
            }
        })
    })

    setTimeout(()=>{
        location.reload();
    },60000)
}