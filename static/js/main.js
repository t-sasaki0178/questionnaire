//フォームの追加
function add_Form(){
    //フォームの原型
    let arche_type = '<div class="q_output"><div class="set"><label class="set_questions"><span class="">第</span><input type="text" name="q_num" class="q_num" value=""><span class="">問</span></label><label class="set_type"><select name="q_select" class="q_select"><option value="singletext">記述問題（１枠・掲載許可なし）</option><option value="doubletext">記述問題（２枠・掲載許可なし）</option><option value="permit_singletext">記述問題（１枠・掲載許可あり）</option><option value="permit_doubletext">記述問題（２枠・掲載許可あり）</option><option value="radio">選択問題（ラジオボタン）</option><option value="radio_plus_text">選択問題（ラジオボタン・記述回答付き）</option><option value="checkbox">選択問題（チェックボックス）</option><option value="matsuda">マツダファンブック固有処理</option><option value="camion">カミオン固有処理</option></select></label></div></div>';
    //フォームの箱
    let base_type = [].slice.call(document.getElementsByClassName('q_output'))[0];
    base_type.insertAdjacentHTML('beforeend',arche_type);
}
let add_form = [].slice.call(document.getElementsByClassName('btn_add'))[0];
add_form.addEventListener('click',function(){add_Form()},false);

//取得ファイル名の表示
function getFiles(){
    //比較用配列
    const checkfiles = ['layout.csv','labeldata.csv','rawdata.csv'];
    //アップファイル名の取得
    const getfiles = Array.from(document.getElementById('files').files).map(file=>file.name);
    //比較用配列と比較し、存在するファイル名には値返却用配列に◯、存在しない場合はXを返す
    let return_checked = ['X','X','X'];
    checkfiles.filter(cf=>{
        getfiles.filter(gf=>{
            if(cf===gf){
                if('layout.csv'===gf){
                    return_checked[0] = '◯';
                }else if('labeldata.csv'===gf){
                    return_checked[1] = '◯';
                }else if('rawdata.csv'===gf){
                    return_checked[2] = '◯';
                }
            }
        });
    });
    //比較結果をフォームのテキスト部分に返す。
    const status = Array.from(document.getElementsByClassName('files_status'));
    let st_c = 0;
    status.filter(st=>{ st.innerHTML = st.innerHTML.replace('X',return_checked[st_c]); st_c++; });
}
