using System;
using System.Net;
using System.Net.Sockets;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Concurrent;
public class SPFixture : IDisposable {
    TcpListener listener; public string BaseUrl;
    public ConcurrentDictionary<string,int> Counts=new ConcurrentDictionary<string,int>();
    public ConcurrentDictionary<string,string> LastRange=new ConcurrentDictionary<string,string>();
    public ConcurrentDictionary<string,string> LastIfRange=new ConcurrentDictionary<string,string>();
    public byte[] Data=Encoding.UTF8.GetBytes("0123456789abcdefghijklmnopqrstuvwxyz");
    public SPFixture() {
        listener=new TcpListener(IPAddress.Loopback,0);listener.Start();
        BaseUrl="http://127.0.0.1:"+((IPEndPoint)listener.LocalEndpoint).Port;
        Task.Factory.StartNew(()=>{while(true){try{var c=listener.AcceptTcpClient();Task.Factory.StartNew(()=>Serve(c));}catch{return;}}});
    }
    void Serve(TcpClient client) {
        using(client) try {
            var stream=client.GetStream();var reader=new StreamReader(stream,Encoding.ASCII,false,1024,true);
            string first=reader.ReadLine();if(first==null)return;
            string path=first.Split(' ')[1],line,range="",ifRange="";
            while(!String.IsNullOrEmpty(line=reader.ReadLine())){if(line.StartsWith("Range:"))range=line.Substring(6).Trim();if(line.StartsWith("If-Range:"))ifRange=line.Substring(9).Trim();}
            LastRange[path]=range;LastIfRange[path]=ifRange;int count=Counts.AddOrUpdate(path,1,(k,n)=>n+1);
            int status=200,offset=0;string extra="",type="application/octet-stream";
            byte[] body=Data;
            if(path=="/page"){type="text/html";body=Encoding.UTF8.GetBytes("<html><title>Local test</title></html>");}
            if(path=="/json"){type="application/json";body=Encoding.UTF8.GetBytes("{\"value\":[{\"Name\":\"fixture.txt\",\"Length\":36}],\"odata.nextLink\":null}");}
            if(path.StartsWith("/_api/web?")){type="application/json";body=Encoding.UTF8.GetBytes("{\"Url\":\"https://example.sharepoint.com/sites/Test\"}");}
            if(path=="/missing")status=404;
            if(path=="/auth")status=401;
            if((path=="/throttle" || path=="/retry-throttle") && count==1){status=429;extra="Retry-After: 0\r\n";}
            if(path=="/html"){type="text/html";body=Encoding.UTF8.GetBytes("<title>Login</title>");}
            if(path=="/slow"){System.Threading.Thread.Sleep(500);}
            if(range!="" && path!="/ignore") {
                offset=int.Parse(range.Substring(6).TrimEnd('-'));status=206;
                extra="Content-Range: bytes "+(path=="/bad-range"?offset+1:offset)+"-"+(Data.Length-1)+"/"+Data.Length+"\r\n";
                body=new byte[Data.Length-offset];Array.Copy(Data,offset,body,0,body.Length);
            }
            string head="HTTP/1.1 "+status+" Test\r\nConnection: close\r\nContent-Type: "+type+"\r\nContent-Length: "+body.Length+"\r\nETag: \"v1\"\r\n"+extra+"\r\n";
            byte[] h=Encoding.ASCII.GetBytes(head);stream.Write(h,0,h.Length);
            if(path=="/broken" && count==1)stream.Write(body,0,10);else stream.Write(body,0,body.Length);
            stream.Flush();
        }catch{}
    }
    public void Dispose(){listener.Stop();}
}
