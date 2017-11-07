//****************************************************************
// Ajout de  méthodes à l'objet Array
// (a) Array.prototype.supprime=function(i) retourne un tableau sans l'indice i
// (b) Array.prototype.echange=function(i,j) retourne un tableau avec un échange entre les indices i et j
// (c) Array.prototype.alea=function(t,min,max)retourne un tableau. La taille du tableau est t.Le contenu des cellules sont des entiers entre min et max
// (d) Array.prototype.max=function() retourne le plus graand nombre d'un tableau.
// (e) Array.prototype.mode=function() retourne le plus nombre le plus souvent présent dans le tableau.
// (f) Array.prototype.visuel=function(canvas) dessine dans la balis d'id canvas un ensemble de points (indice;tab[indice]) qui représente les éléments du tableau.
//****************************************************************               

Array.prototype.supprime=function(ind) {
    for(var i=0;i<this.length-1;i++){
        if(i>=ind){
            this[i]=this[i+1]
        }
    }
    this.length--
    return this
}
Array.prototype.echange=function(i,j){
    var t=this[i]
    this[i]=this[j]
    this[j]=t
    return this
}
Array.prototype.alea=function(taille,min,max){
    for (var i=0; i<taille; i++) {
        this[i]=min+Math.floor(Math.random()*(max-min+1))
    }
    return this
}
Array.prototype.max=function(array){
    var max=this[0];
    var tI;
    for (var i=0; i<this.length; i++) {
        if(this[i]>max){
            max=this[i];
            tI = i;
        }
    }
    if (array) {
        return [max,tI];
    }
    else{
        return max;  
    }

}

Array.prototype.mode=function(){
    var aux=[{car:this[0],nb:1}]
    var trouve=false
    for (var i=1; i<this.length; i++) {
        for (var j=0;j<aux.length;j++){
            if(this[i]==aux[j].car){
                aux[j].nb++
                trouve=true
            }
        }
        if(!trouve){
            aux.push({car:this[i],nb:1})
        }
        trouve=false
    }
    var valMax=[]
    var maximum=aux[0].nb
    for (var k=1; k<aux.length; k++) {
        if(aux[k].nb>maximum){maximum=aux[k].nb}
    }
    for (var l=0; l<aux.length; l++) {
        if(aux[l].nb==maximum){valMax.push(aux[l].car)}
    }
    return {valMAX: valMax, nbREP: maximum}
}

Array.prototype.visuel=function(canvas){
    var ctx=document.getElementById(canvas).getContext('2d')
    var w=document.getElementById(canvas).width
    var h=document.getElementById(canvas).height
    ctx.clearRect(0,0,w,h);
    var qx=w/this.length;
    var qy=h/this.max(false);         
    for (var i=0; i<this.length; i++) {
        ctx.beginPath()
        ctx.arc(qx/2+i*qx,qy/2+ this[i]*qy,1,0,2*Math.PI,true)
        ctx.stroke()
    }

}

Array.prototype.triSelection = function(){
    var array = [];
    for(k=0;k<this.length;k++){
        array.push(this[k]);
    }
    for(var k = 0; k<this.length;k++){
        var max = array.max(true);
        array.splice(max[1],1);
        for(let l=0;l<this.length;l++){
            if(max[0] == this[l]){
                this.echange(l,this.length-1-k);
            }
        }
    }
    return this;
}

Array.prototype.triBulles = function(){
    while(!//est triÈ){
          for(let k=1;k<this.length+1;k++){
        if(this[k-1]<this[k]){
            this.echange(k-1,k);
        }
    }
}
return this;
}

Array.prototype.quickSort = function(){}




