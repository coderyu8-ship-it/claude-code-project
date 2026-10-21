あなたは Shopify のコードエディタ（Chrome で開いているタブ）を操作して、カートページを修正します。
以下の条件・手順を厳守してください。

## 絶対条件
- 編集してよいのはテーマ「【t.style検証】カート画面編集」だけ。公開中（ライブ）のテーマは開かない・触らない。
- テーマの「公開」ボタンは絶対に押さない。
- 編集するのは手順1で特定した sections 内のカート用ファイル 1 つだけ。他のファイル（locales / layout / config / snippets / assets）は変更しない。
- 想定と違う構造だった場合（探しているコードが見つからない、複数ある等）は、推測で編集せず、その時点で作業を止めて状況を報告する。
- 保存（Ctrl+S）の前に、変更前後の該当箇所をスクリーンショットで報告する。

## 手順1：編集対象ファイルを特定
1. 左の検索（Ctrl+Shift+F）で `熨斗を追加する` を全ファイル検索する。
2. ヒットした sections/ 配下のファイルを開く（例：sections/main-cart.liquid）。
3. ヒットしない場合は templates/cart.json を開き、main セクションの "type" の値から sections/<type>.liquid を開く。
4. ファイル名を報告する。

## 手順2：差し替え A（のし）
以下の 2 行（空白・改行の差は許容）を探し、削除して、その位置に【コードA】を貼り付ける。

```
<a class="btn" href="/products/noshi" style="margin-bottom:1rem">熨斗を追加する</a>
<div style="margin-bottom: 2rem;">のしの表書き、名入れの印刷をご希望の方は、下記備考欄に記載をお願いいたします。</div>
```

【コードA】
```liquid
{%- assign a = cart.attributes -%}
<div class="cart-gift">
  <p class="cart-gift__ttl">のし・名入れ（無料）</p>

  <label for="CartNoshi">のし</label>
  <select name="attributes[のし]" id="CartNoshi" class="input-full">
    {%- for opt in "不要,紅白結び切り（慶事）,黒白結び切り（弔事）" | split: "," -%}
      <option value="{{ opt }}" {% if a['のし'] == opt %}selected{% endif %}>{{ opt }}</option>
    {%- endfor -%}
  </select>

  <label for="CartOmote">表書き</label>
  <select name="attributes[表書き]" id="CartOmote" class="input-full">
    {%- for opt in "御礼,内祝,御祝,御中元,御歳暮,御年賀,粗品,志,御供,その他（備考欄に記入）" | split: "," -%}
      <option value="{{ opt }}" {% if a['表書き'] == opt %}selected{% endif %}>{{ opt }}</option>
    {%- endfor -%}
  </select>

  <label for="CartNaire">名入れ（例：田丸）</label>
  <input type="text" name="attributes[名入れ]" id="CartNaire" class="input-full"
         value="{{ a['名入れ'] }}" maxlength="20" placeholder="空欄の場合は名入れなし">
</div>
```

## 手順3：差し替え B（送料）
購入手続きボタン（name="checkout"）の下にある、「送料、税金、割引価格はチェックアウト時に計算されます」を出力している `<small> ～ </small>` ブロックを探す。
ファイル内では翻訳キー（例：`{{ 'cart.general.shipping_at_checkout' | t }}` など shipping を含むもの）で書かれている可能性が高い。ファイル内検索で `shipping` を探して特定する。
その `<small> ～ </small>` を丸ごと削除し（外側の `<div class="cart__item-row text-center">` は残す）、その位置に【コードB】を貼り付ける。

【コードB】
```liquid
<div class="cart-ship" id="CartShip">
  <p class="cart-ship__ttl">送料を確認する</p>
  <div class="cart-ship__row">
    <select id="CartShipPref" class="input-full" aria-label="都道府県">
      <option value="">都道府県</option>
      {%- assign prefs = "JP-01:北海道,JP-02:青森県,JP-03:岩手県,JP-04:宮城県,JP-05:秋田県,JP-06:山形県,JP-07:福島県,JP-08:茨城県,JP-09:栃木県,JP-10:群馬県,JP-11:埼玉県,JP-12:千葉県,JP-13:東京都,JP-14:神奈川県,JP-15:新潟県,JP-16:富山県,JP-17:石川県,JP-18:福井県,JP-19:山梨県,JP-20:長野県,JP-21:岐阜県,JP-22:静岡県,JP-23:愛知県,JP-24:三重県,JP-25:滋賀県,JP-26:京都府,JP-27:大阪府,JP-28:兵庫県,JP-29:奈良県,JP-30:和歌山県,JP-31:鳥取県,JP-32:島根県,JP-33:岡山県,JP-34:広島県,JP-35:山口県,JP-36:徳島県,JP-37:香川県,JP-38:愛媛県,JP-39:高知県,JP-40:福岡県,JP-41:佐賀県,JP-42:長崎県,JP-43:熊本県,JP-44:大分県,JP-45:宮崎県,JP-46:鹿児島県,JP-47:沖縄県" | split: "," -%}
      {%- for p in prefs -%}{%- assign kv = p | split: ":" -%}
        <option value="{{ kv[0] }}">{{ kv[1] }}</option>
      {%- endfor -%}
    </select>
    <input type="text" id="CartShipZip" class="input-full" inputmode="numeric"
           placeholder="郵便番号（例：377-0102）" aria-label="郵便番号">
    <button type="button" class="btn" id="CartShipBtn">計算</button>
  </div>
  <p class="cart-ship__result" id="CartShipResult" aria-live="polite"></p>
  <small>お届け先ごとの送料です。税込。最終金額はご注文手続き画面でご確認ください。</small>
</div>

<script>
(function(){
  var btn=document.getElementById('CartShipBtn'), out=document.getElementById('CartShipResult');
  btn.addEventListener('click', function(){
    var pref=document.getElementById('CartShipPref').value,
        zip=document.getElementById('CartShipZip').value.trim();
    if(!pref||!zip){ out.textContent='都道府県と郵便番号を入力してください'; return; }
    out.textContent='計算中…';
    var q='shipping_address[country]=Japan&shipping_address[province]='+encodeURIComponent(pref)
         +'&shipping_address[zip]='+encodeURIComponent(zip);
    fetch('/cart/shipping_rates.json?'+q,{credentials:'same-origin'})
      .then(function(r){ return r.json(); })
      .then(function(d){
        if(!d.shipping_rates){ out.textContent='郵便番号をご確認ください'; return; }
        if(!d.shipping_rates.length){ out.textContent='この地域への配送はご注文手続き画面でご確認ください'; return; }
        out.innerHTML=d.shipping_rates.map(function(s){
          return s.name+'：<strong>¥'+Number(s.price).toLocaleString()+'</strong>';
        }).join('<br>');
      })
      .catch(function(){ out.textContent='送料を取得できませんでした。ご注文手続き画面でご確認ください'; });
  });
})();
</script>
```

## 手順4：CSS 追加
同じファイルの一番最後（ファイル末尾。`{% schema %}` がある場合はその直前）に【コードC】を貼り付ける。

【コードC】
```html
<style>
.cart-gift{margin-bottom:2rem;padding:1.2rem;border:1px solid #e5dccb;background:#F9F5ED}
.cart-gift__ttl{font-weight:700;margin-bottom:.8rem}
.cart-gift label{display:block;margin:.8rem 0 .3rem;font-size:.9em}
.cart-ship{margin:1rem 0;padding:1rem;border:1px solid #e5dccb;text-align:left}
.cart-ship__ttl{font-weight:700;margin-bottom:.5rem}
.cart-ship__row{display:grid;grid-template-columns:1fr 1fr auto;gap:.5rem}
.cart-ship__result{margin:.6rem 0 .3rem;font-size:1.05em}
@media (max-width:480px){.cart-ship__row{grid-template-columns:1fr}}
</style>
```

## 手順5：貼り付け時の注意
- エディタは自動インデント・括弧の自動補完が働く。貼り付けはキー入力ではなくクリップボード貼り付け（Ctrl+V）で行い、貼り付け後に括弧や `{% %}` が二重になっていないか目視確認する。
- 貼り付け後、エディタ下部のエラー表示（×や⚠の数）が増えていないか確認する。増えた場合は保存せず報告。

## 手順6：保存と確認
1. 変更前後のスクリーンショットを報告し、Ctrl+S で保存する。
2. 保存エラー（Liquid 構文エラー等）が出たら、内容をそのまま報告して止まる。
3. テーマ一覧から「【t.style検証】カート画面編集」の「プレビュー」を開き、商品を 1 つカートに入れて /cart を表示する。
4. 次を確認して結果を報告する：
   - のし・表書き・名入れの入力欄が表示される
   - 都道府県「東京都」、郵便番号「100-0001」で「計算」を押すと送料が表示される
   - ウィンドウ幅を狭めても表示が崩れない
5. テスト注文・決済は行わない。テーマの公開もしない。
