using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading.Tasks;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using System.Xml.Serialization;


namespace VSBouyomi {

	public class EventHolder
	{
		public EnvDTE.BuildEvents BuildEvents { get; set; }
		public EnvDTE.SolutionEvents SolutionEvents { get; set; }
		public EnvDTE.DTEEvents DteEvents { get; set; }
		public EnvDTE.DebuggerEvents DebuggerEvents { get; set; }
	}


	public class BouyomiTcpClient {
		private const string DefaultHost = "127.0.0.1";
		private const int DefaultPort = 50001;

		public static string TargetHost { get; set; }
		public static int TargetPort { get;set; }


		const byte Code = 0;
		const Int16 Voice = 1;
		const Int16 Volume = -1;
		const Int16 Speed = -1;
		const Int16 Tone = -1;

		/// <summary>
		/// 引数2のコンストラクタ
		/// </summary>
		/// <param name="host">ホストIP</param>
		/// <param name="port">ポート番号</param>
		public BouyomiTcpClient(string host,int port) {
			TargetHost = host;
			TargetPort = port;
		}

		/// <summary>
		/// IPエンドポイントを指定したコンストラクタ
		/// </summary>
		/// <param name="endPoint">エンドポイント</param>
		public BouyomiTcpClient(IPEndPoint endPoint) :this(endPoint.Address.ToString(),endPoint.Port) {}

		/// <summary>
		/// 引数0のコンストラクタ　Loopbackの50001番に接続
		/// </summary>
		public BouyomiTcpClient():this(DefaultHost,DefaultPort) {}

		/// <summary>
		/// 発言させる
		/// </summary>
		/// <param name="Text">文章</param>
		public void Talk(string Text) {
			Talk(Text, Speed, Tone, Volume, Voice, Code, true);
		}

		public void Talk(string Text,bool force) {
			Talk(Text, Speed, Tone, Volume, Voice, Code, force);
		}

		public async void Talk(string Text, short speed, short tone, short volume, short voice, byte code,bool force) {
			bool status;
			if (!force) do {
					status = await GetNowPlaying();
				} while (status);
			else {
				Clear();
			}
			Talk(Text, speed, tone, volume, voice, code);
		}

		/// <summary>
		/// 発言させる
		/// </summary>
		/// <param name="Text"></param>
		/// <param name="speed"></param>
		/// <param name="tone"></param>
		/// <param name="volume"></param>
		/// <param name="voice"></param>
		/// <param name="code"></param>
		public void Talk(string Text,short speed,short tone,short volume,short voice,byte code) {
			Text = Text ?? "";
			var mes = Encoding.UTF8.GetBytes(Text);
			Int32 length = mes.Length;
			Post(bw => {
				bw.Write((short)BouyomiCommand.Talk);
				bw.Write(speed);
				bw.Write(tone);
				bw.Write(volume);
				bw.Write(voice);
				bw.Write(code);
				bw.Write(length);
				bw.Write(mes);
			});
		}

		/// <summary>
		/// 一時停止
		/// </summary>
		public void Pause() {
			Post(bw => {
				bw.Write((Int16)BouyomiCommand.Pause);
			});
		}

		/// <summary>
		/// 再開
		/// </summary>
		public void Resume() {
			Post(bw => {
				bw.Write((Int16)BouyomiCommand.Resume);
			});
		}

		/// <summary>
		/// スキップ
		/// </summary>
		public void Skip() {
			Post(bw => bw.Write((Int16)BouyomiCommand.Skip));
		}

		/// <summary>
		/// タスクを消す
		/// </summary>
		public void Clear() {
			Post(bw => bw.Write((Int16)BouyomiCommand.Clear));
		}

		/// <summary>
		///	ポーズ中かどうか
		/// </summary>
		/// <returns></returns>
		public async Task<bool> GetPause() {
			var gets = await Request(BouyomiCommand.GetPause);
			return gets[0] == 0x01;
		}

		/// <summary>
		/// 再生中かどうか
		/// </summary>
		/// <returns></returns>
		public async Task<bool> GetNowPlaying() {
			var gets = await Request(BouyomiCommand.GetNowPlaying);
			return gets[0] == 0x01;
		}

		/// <summary>
		/// 残りタスク数を返します
		/// </summary>
		/// <returns></returns>
		public async Task<int> GetTaskCount() {
			var gets = await Request(BouyomiCommand.GetTaskCount);
			int rt = 0;
			for (int i = 0; i < 4; i++) {
				rt += (rt << 4) + gets[i];
			}
			return rt;
		}

		/// <summary>
		/// ホストへ接続してActionを実行
		/// </summary>
		/// <param name="action"></param>
		private async void Post(Action<BinaryWriter> action) {
			TcpClient tcpClient = new TcpClient();
			await tcpClient.ConnectAsync(TargetHost, TargetPort);
			using (var ns = tcpClient.GetStream())
			using (var bw = new BinaryWriter(ns)) {
				action(bw);
			}
			tcpClient.Close();
		}

		/// <summary>
		/// レスポンスのあるリクエストを送る
		/// </summary>
		/// <param name="bouyomiCommand"></param>
		/// <returns></returns>
		private async Task<byte[]> Request(BouyomiCommand bouyomiCommand) {
			using(var tcp=new TcpClient()) {
				await tcp.ConnectAsync(TargetHost, TargetPort);
				using(var stream = tcp.GetStream()) {
					var bw = new BinaryWriter(stream);
					var br = new BinaryReader(stream);

					//Write Command;
					bw.Write((Int16)bouyomiCommand);
					bw.Flush();

					//Read Response
					var gets = br.ReadBytes(bouyomiCommand == BouyomiCommand.GetTaskCount ? 4 : 1);
					bw.Close();
					br.Close();
					return gets;
				}
			}
		}

		public enum BouyomiCommand : Int16 {
			Talk = 0x0001,
			Pause = 0x0010,
			Resume = 0x0020,
			Skip = 0x0030,
			Clear = 0x0040,
			GetPause = 0x0110,
			GetNowPlaying = 0x0120,
			GetTaskCount = 0x0130
		}
	}
}
