# IITC-sirius-
出来る事
●IITC を利用しての ポータル仕様 抽出(P8になりそうなPortal and レア度の高いMOD入りPortal)
　必須：なし
　便利：show list of portals Plugin (画面内Portalを表形式で表示)
●IITC を利用しての オーナーポータル表示(Portalオーナー and レゾオーナー)
　必須：highlight portals by level color Plugin (Portalをレベルで色分け)
　便利：show list of portals Plugin (画面内Portalを表形式で表示)
　　　　Portal Level Numbers Plugin (Portalレベルを数字で表示)

使い方
●ソース内の以下の部分で動作を変更可能(変更後 IITCページを更新の事)
--以下--
// CONFIG OPTIONS ////////////////////////////////////////////////////
//@@@ここで設定可能
window.siriussk8er = 3; // 1=@1Portal探し 2=myPortal表示 3=両方
//@@@ここまで
--ここまで--
