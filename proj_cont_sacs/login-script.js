function click_input(){
    var input = document.getElementById('loja')
    input.value=""
    input.style.color="black"
}

function mouseout_input(){
    var input = document.querySelector('input#loja')
    if (input.value == "") {
        input.value = "Ex: Guaxupé"
        input.style.color="rgb(172, 172, 172)"
    }
}

function click_senha(){
    var input = document.getElementById('senha')
    input.type="password"
    input.value=""
    input.style.color="black"
}

function mouseout_senha(){
    var input = document.querySelector('input#senha')
    if (input.value == ""){
    input.type="text"
    input.value = "Ex: 123"
    input.style.color="rgb(172, 172, 172)"
    }
}