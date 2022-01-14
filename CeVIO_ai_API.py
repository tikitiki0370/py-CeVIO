from re import split as sp

import win32com.client



class StartupError(Exception):
    def __init__(self,number):
        self.error_number = int(number)
        self.error_text={
            -1:"インストール状態が不明です",
            -2:"ファイルが見つかりません",
            -3:"プロセスの起動に失敗しました",
            -4:"エラーにより終了しました"
        }
    def __str__(self) -> str:
        return (f"CeVIO_AIの起動に失敗しました:{self.error_text[self.error_number]}")

class CeVIOai:

    __talker_name= []
    __emotion = {}

    def __init__(self):
        #CeVIO起動API
        CeVIOai.service = win32com.client.Dispatch("CeVIO.Talk.RemoteService2.ServiceControl2")
        service_status = CeVIOai.service.StartHost(False)
        if service_status < 0:
            raise StartupError(service_status)

        #API接続
        self.__talker = win32com.client.Dispatch("CeVIO.Talk.RemoteService2.Talker2")

        #話し手取得
        if not CeVIOai.__talker_name:
            self.get_talker()

        #初期設定
        self.__talker.Cast = CeVIOai.__talker_name[0]
        self.__talker.Volume = 50

    def _list_check(self, text:str):
        #list以外を変換
        list_text = []
        if type(text) != list:
            list_text.append(str(text))
        else:
            list_text += text
        return list_text


    def generate(self,text:list = "", path:str = "./output_*.wav"):
        """
        wav書き出し
        """
        return_text = []
        text = self._list_check(text)
        for i, speak in enumerate(text):
            if "*" in path:
                temp = path.replace("*", i)
            return_text.append(self.__talker.OutputWaveToFile(speak,temp))
        return return_text

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
        speak_text = self._list_check(text)
        #読み上げ
        for speak in speak_text:
            #200文字以上は自動分割
            if len(speak) >= 200:
                return_text += self.speak(self.split_speak_text(speak))
                continue
            status = self.__talker.Speak(speak)
            status.Wait_2(wait_time)
            return_text.append(status.IsSucceeded)
        return return_text

    def stop(self):
        """
        読み上げの停止
        """
        return self.__talker.Stop()


    def split_speak_text(self,text:str,
                        pattern:str =r"\s|\_|\\|\(|\)|\"|\'|\.|\,|、|。|「|」") -> list:
        """
        文章を切り分ける

        Parametars
        ----------
        text:
            切り分ける文
        pattern:
            区切り文字の正規表現
        """
        return_text = []
        temp = sp(pattern, text)
        for i in temp:
            if len(i) == 200:
                return ["切り分けられませんでした"]
            elif len(i) == 0:
                continue
            return_text.append(i)
        return return_text


    def reset_emotion(self, skip:list = []):
        """
        感情値を初期化
        感情値を0にします

        Parametars
        ----------
        skip:
            操作しない感情
        """
        emo_list = self.get_select_emotion(self.get_cast())
        for i in emo_list:
                if i in skip:
                    continue
                self.__talker.Components.ByName(i).Value = 0
        return "正常に変更されました"


    def change_emotion(self, value:dict, mode:bool = True):
        """
        感情パラメータの変更

        Parameters
        -----------
        value:
            {感情の名前(str) : 感情値(int)}
        mode:
            True  -> 選択しなかった感情値を0にする
            False -> 選択しなかったの感情値を操作しない
        """
        temp = value.keys()
        #選択された感情値を設定
        for i in temp:
            if i in  CeVIOai.__emotion[self.__talker.Cast]:
                self.__talker.Components.ByName(i).Value = value[i]
            else:
                return f"{i}は存在しません"
        #他の感情値を0に
        if mode == True:
            return self.reset_emotion(temp)


    def set_emotion(self, value:str):
        """
        感情パラメータの変更
        選択した感情の値を100にし、他のパラメータを0にします

        Parameters
        ----------
        value:
            感情の名前
        """
        #選択された感情値を100に
        if value in  CeVIOai.__emotion[self.__talker.Cast]:
            self.__talker.Components.ByName(value).Value = 100
        else:
            return f"{value}は存在しません"
        #他の感情値を0に
        self.reset_emotion([value])
        return f"感情パラメーターを{value}に切り替えました"


    def set_talker(self, name:str):
        """
        話し手の変更
        """
        if name in CeVIOai.__talker_name:
            self.__talker.Cast = name
            self.__talker.Components.ByName(CeVIOai.__emotion[name][0]).name = 100
        else:
            return f"{name}は存在しません"

    def set_tone(self, value:int = 50):
        """
        声の高さを設定
        """
        if value <=100 and value >=0:
            self.__talker.Tone = value 
            return f"{value}変更されました" if value == 50 else "リセットされました"
        else:
            return "値が不正です"

    def set_speed(self, value:int = 50):
        """
        読み上げ速度を設定
        """
        if value <= 100 and value >=0:
            self.__talker.Speed = value
            return f"{value}変更されました" if value == 50 else "リセットされました"
        else:
            return "値が不正です"

    # def set_tonescale(self, value:int = 50):
    #     """
    #     抑揚を設定
    #     """
    #     if value <= 100 and value >=0:
    #         self.__talker.ToneScale = value
    #         return f"{value}に変更されました" if value == 50 else "リセットされました"
    #     else:
    #         return "値が不正です"

    def set_alpha(self, value:int = 50):
        """
        声質を設定
        """
        if value <=100 and value >=0:
            self.__talker.Alpha = value
            return f"{value}に変更されました" if value == 50 else "リセットされました"
        else:
            return "値が不正です"

    def set_volume(self, value:int = 50):
        """
        音量を設定
        """
        if value <=100 and value >=0:
            self.__talker.Volume = value
            return f"{value}変更されました" if value == 50 else "リセットされました"
        else:
            return "値が不正です"


    def get_talker(self) -> None:
        """
        話し手の設定を更新
        """

        #話し手一覧を取得
        CeVIOai.__talker_name = [self.__talker.AvailableCasts.At(i) for i in range(self.__talker.AvailableCasts.Length)]

        #現在の話し手を保存
        try:temp = self.__talker.Cast
        except: pass

        #話し手を変更し全員の感情値を取得
        for i in CeVIOai.__talker_name:
            self.__talker.Cast = i
            CeVIOai.__emotion[i] = [self.__talker.Components.At(i).Name for i in range(self.__talker.Components.Length)]

        #元の話し手の復元
        try:self.__talker.Cast = temp
        except: pass


    # def get_parameters(self) -> list:
    #     return_text = 
    #     return return_text

    def get_tone(self) -> int:
        return self.__talker.Tone

    def get_speed(self) -> int:
        return self.__talker.Speed

    # def get_tonescale(self) -> int:
    #     return self.__talker.LogF0Scale

    def get_alpha(self) -> int:
        return self.__talker.Alpha

    def get_volume(self) -> int:
        return self.__talker.Volume

    def get_text_duration(self,text:str) -> float:
        """
        セリフの長さを取得
        
        Parameters
        ----------
        text:
            セリフ
        """
        return self.__talker.GetTextDuration(text)

    def get_talkername(self) -> list:
        """
        設定できる話し手の一覧を取得
        """
        return CeVIOai.__talker_name

    def get_emotion(self) -> dict:
        """
        話し手ごとの感情値を取得

        Return
        --------
        {話し手の名前(str) : 感情値(int)}
        """
        return CeVIOai.__emotion

    def get_select_emotion(self,name:str) -> list:
        """
        選択した話し手の感情値を取得
        """
        return CeVIOai.__emotion[name]

    def get_cast(self) -> str:
        """
        現在設定されている話し手を取得
        """
        return self.__talker.Cast

    def get_emotion_value(self) -> dict:
        """
        現在の感情値を取得
        """
        return {
            i : self.__talker.Components.ByName(i).Value 
            for i in self.get_select_emotion(self.get_cast())
        }


if __name__ == "__main__":
    test_meg = "これはテストメッセージです"
    test = CeVIOai()
    print(test.get_emotion())
    print(test.speak(test_meg))
