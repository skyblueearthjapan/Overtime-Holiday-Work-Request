// ====== フォームの質問アイテム取得（タイトルで検索） ======

function findItemByTitle_(form, title) {
  const items = form.getItems();
  for (const it of items) {
    if (normalize_(it.getTitle()) === title) return it;
  }
  throw new Error(`フォーム内に質問が見つかりません: "${title}"（テンプレに同名の質問を作ってください）`);
}

function setDropdownChoices_(item, choices) {
  // DropdownItem or ListItem対応（テンプレの形式に合わせる）
  const type = item.getType();
  if (type === FormApp.ItemType.LIST) {
    const li = item.asListItem();
    li.setChoiceValues(choices);
    return;
  }
  if (type === FormApp.ItemType.MULTIPLE_CHOICE) {
    const mi = item.asMultipleChoiceItem();
    mi.setChoiceValues(choices);
    return;
  }
  if (type === FormApp.ItemType.DROP_DOWN) {
    // ※DROP_DOWN は ItemType に存在しない場合があるため保険（環境差）
    item.asListItem().setChoiceValues(choices);
    return;
  }
  throw new Error(`この質問はプルダウン/ラジオではありません: "${item.getTitle()}" type=${type}`);
}
