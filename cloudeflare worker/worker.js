export default {
  async fetch(request, env, ctx) {
    const GAS_EXEC_URL = "url";
    
    // LINE等の検証アクセス(GET)には即座に200を返して終了
    if (request.method !== "POST") {
      return new Response("OK", { status: 200 });
    }

    try {
      // LINEからのメッセージ内容(JSON)を取得
      const body = await request.text();

      // 裏側でGASに投げ込む（完了を待たずにLINEにはOKを返す：非同期実行）
      ctx.waitUntil(
        fetch(GAS_EXEC_URL, {
          method: "POST",
          headers: { 
            "Content-Type": "application/json" 
          },
          body: body,
          redirect: "follow", // GASのリダイレクト(302)を自動追従
        }).then(res => {
          // 成功確認用ログ（Cloudflareのログで確認可能）
          console.log(`Sent to GAS. Status: ${res.status}`);
        }).catch(err => {
          // エラーログ
          console.log(`Failed to send to GAS: ${err}`);
        })
      );

      // LINEサーバーには即座に成功(200)を返す
      return new Response("OK", { status: 200 });

    } catch (e) {
      // 万が一のWorker内エラー
      console.log(`Worker Error: ${e}`);
      return new Response("Error", { status: 500 });
    }
  },
};
