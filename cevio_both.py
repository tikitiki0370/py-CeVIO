import win32com.client

class StartupError(Exception):
    def __init__(self,number, name):
        self.error_number = int(number)
        self.name = name
        self.error_text={
            -1:"インストール状態が不明です",
            -2:"ファイルが見つかりません",
            -3:"プロセスの起動に失敗しました",
            -4:"エラーにより終了しました"
        }
    def __str__(self) -> str:
        return (f"CeVIO_{self.name}の起動に失敗しました:{self.error_text[self.error_number]}")

class CeVIOboth:
    service = None
    __talker_ai = None
    __talker_cs = None
    __talker_name_ai = ["小春六花"]
    __talker_name_cs = ["さとうささら"]
    def __init__(self):

        #CeVIO_AI起動API
        if not CeVIOboth.service:
            CeVIOboth.service = win32com.client.Dispatch("CeVIO.Talk.RemoteService2.ServiceControl2")
        service_status = CeVIOboth.service.StartHost(False)
        if service_status < 0:
            raise StartupError(service_status, "AI")

        if not CeVIOboth.service:
            CeVIOboth.service = win32com.client.Dispatch("CeVIO.Talk.RemoteService.ServiceControl")
        service_status = CeVIOboth.service.StartHost(False)
        if service_status < 0:
            raise StartupError(service_status)

        #CeVIO_CS起動API
        if not CeVIOboth.service:
            CeVIOboth.service = win32com.client.Dispatch("CeVIO.Talk.RemoteService.ServiceControl")
        service_status = CeVIOboth.service.StartHost(False)
        if service_status < 0:
            raise StartupError(service_status, "CS")

        #AI_API接続
        if not CeVIOboth.__talker_ai:
            CeVIOboth.__talker_ai = win32com.client.Dispatch("CeVIO.Talk.RemoteService2.Talker2")

        #CS_API接続
        if not CeVIOboth.__talker_cs:
            CeVIOboth.__talker_cs = win32com.client.Dispatch("CeVIO.Talk.RemoteService2.Talker2")

        # #話し手取得
        # if not CeVIOboth.__talker_name:
        #     self.get_talker()

        #初期設定
            if not CeVIOboth.__talker_ai.Cast:
                CeVIOboth.__talker_ai.Cast = CeVIOboth.__talker_name[0]
                CeVIOboth.__talker_ai.Volume = 50

        #初期設定
            if not CeVIOboth.__talker_cs.Cast:
                CeVIOboth.__talker_cs.Cast = CeVIOboth.__talker_name[0]
                CeVIOboth.__talker_cs.Volume = 50


    def speak(self,text:list, wait_time:float = -1):
        """
        読み上げを開始

        Parameters
        -----------
        text:
            読み上げる文章
        wait_time:
            再生終了までの最大待機時間
        """
        return_text = []
        #リストに変更
        # speak_text = CeVIOboth._list_check(text)
        #読み上げ
        for speak in text:
            #500文字以上は自動分割
            # if len(speak) >= 500:
            #     return_text += self.speak(self.split_speak_text(speak))
            #     continue
            status = CeVIOboth.__talker_ai.Speak(speak)
            status.Wait_2(wait_time)
            return_text.append(status.IsSucceeded)
        return return_text


if __name__ == "__main__":
    test_meg = "これはテストメッセージです"
    test = CeVIOboth()
    rest = CeVIOboth()
    print(test.speak(test_meg))
    print(rest.speak(test_meg))