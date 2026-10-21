using System;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Concurrent;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;

namespace SharePointPS {
    public static class LocalPaths {
        static LocalPaths() {
            AppContext.SetSwitch("Switch.System.IO.UseLegacyPathHandling",false);
            AppContext.SetSwitch("Switch.System.IO.BlockLongPaths",false);
        }
        public static string Native(string path) {
            if(path.StartsWith(@"\\?\")) return path;
            string full=Path.GetFullPath(path);
            if(full.StartsWith(@"\\")) return @"\\?\UNC\"+full.Substring(2);
            return @"\\?\"+full;
        }
        public static bool Exists(string path) { string p=Native(path); return File.Exists(p)||Directory.Exists(p); }
        public static void EnsureDirectory(string path) { Directory.CreateDirectory(Native(path)); }
    }
    [DataContract] public class Partial {
        [DataMember(Name="identity")] public string Identity;
        [DataMember(Name="size")] public long Size;
        [DataMember(Name="etag")] public string ETag;
    }
    public class Result {
        public bool Success, Authentication, SizeMismatch;
        public long Expected, Received, ResponseLength;
        public string ResponseETag, ContentEncoding;
        public int Status;
        public double RetryAfter;
        public string Message, Route, Correlation;
    }
    public class Transfer {
        public volatile bool Cancel;
        public long Bytes;
        public Task<Result> Task;
        public ConcurrentQueue<string> Events = new ConcurrentQueue<string>();
        public void Start(string[] urls, string target, long size, string identity, CookieContainer cookies) {
            Task = System.Threading.Tasks.Task.Factory.StartNew(() => Run(urls,target,size,identity,cookies));
        }
        static Partial Read(string path) {
            using(var s=File.OpenRead(path)) return (Partial)new DataContractJsonSerializer(typeof(Partial)).ReadObject(s);
        }
        static void Save(string path, Partial value) {
            string temp=path+".tmp";
            using(var s=File.Create(temp)) new DataContractJsonSerializer(typeof(Partial)).WriteObject(s,value);
            if(File.Exists(path)) File.Replace(temp,path,null); else File.Move(temp,path);
        }
        public static string Clean(string value) {
            string t=value ?? "";
            t=Regex.Replace(t,@"(?i)\b(?:authorization|cookie|set-cookie|access_token|refresh_token|id_token|client_secret|password|senha|sig|token)\b[""']?\s*[:=]\s*(?:""[^""]*""|'[^']*'|[^\s;,]+)","[credencial-removida]");
            t=Regex.Replace(t,@"(?i)https?://[^\s<>""']+","[url-removida]");
            t=Regex.Replace(t,@"(?i)Bearer\s+\S+","Bearer [removido]");
            t=Regex.Replace(t,@"\beyJ[A-Za-z0-9_-]+(?:\.[A-Za-z0-9_-]+){1,2}\b","[token-removido]");
            t=Regex.Replace(t,@"[A-Za-z0-9_+/=\-]{64,}","[valor-opaco-removido]");
            t=Regex.Replace(t,@"\b[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}\b","[email-removido]");
            t=Regex.Replace(t,@"[\x00-\x1f\x7f]"," ");
            return t.Substring(0,Math.Min(1200,t.Length));
        }
        static double Delay(string value) {
            double n; DateTimeOffset dt;
            if(double.TryParse(value,System.Globalization.NumberStyles.Float,System.Globalization.CultureInfo.InvariantCulture,out n)) return Math.Max(0,n);
            if(DateTimeOffset.TryParse(value,out dt)) return Math.Max(0,(dt-DateTimeOffset.UtcNow).TotalSeconds);
            return 3;
        }
        static Result HttpError(HttpWebResponse response) {
            var r=new Result {Status=(int)response.StatusCode, RetryAfter=Delay(response.Headers["Retry-After"])};
            r.Authentication=r.Status==401 || r.Status==403 || (r.Status>=300 && r.Status<400);
            Guid id; if(Guid.TryParse(response.Headers["SPRequestGuid"],out id)) r.Correlation=id.ToString();
            // Read a bounded body, but never expose raw payload, innererror or headers.
            string code="";
            try {
                using(var s=response.GetResponseStream()) {
                    byte[] b=new byte[16384]; int n=0,k;
                    while(n<b.Length && (k=s.Read(b,n,b.Length-n))>0) n+=k;
                    string text=Encoding.UTF8.GetString(b,0,n);
                    if((response.ContentType??"").Contains("json")) {
                        using(var json=JsonReaderWriterFactory.CreateJsonReader(b,0,n,Encoding.UTF8,System.Xml.XmlDictionaryReaderQuotas.Max,null)) {
                            var doc=new System.Xml.XmlDocument();doc.XmlResolver=null;doc.Load(json);
                            var node=doc.SelectSingleNode("/root/error/code | /root/odata.error/code");
                            if(node!=null) code=" codigo="+Clean(node.InnerText);
                            node=doc.SelectSingleNode("/root/error/message/value | /root/odata.error/message/value");
                            if(node==null) node=doc.SelectSingleNode("/root/error/message[@type='string'] | /root/odata.error/message[@type='string']");
                            if(node!=null) code+=" mensagem="+Clean(node.InnerText);
                        }
                    } else if((response.ContentType??"").Contains("html")) {
                        var m=Regex.Match(text,@"<title\b[^>]*>(.*?)</title>",RegexOptions.IgnoreCase|RegexOptions.Singleline);
                        if(m.Success) code=" titulo="+Clean(Regex.Replace(m.Groups[1].Value,"<[^>]+>",""));
                    }
                }
            } catch { }
            r.Message="HTTP "+r.Status+code;
            return r;
        }
        Result Run(string[] urls,string target,long size,string identity,CookieContainer cookies) {
            try {
                for(int route=0;route<urls.Length;route++) {
                    Result r=One(urls[route],target,size,identity,cookies);
                    r.Route=route.ToString();
                    if((r.Status==400 || r.Status==404) && route+1<urls.Length) { Events.Enqueue("ROTA_ALTERNATIVA http="+r.Status); continue; }
                    return r;
                }
            } catch(OperationCanceledException) { return new Result {Message="Cancelado; parcial preservado."}; }
              catch(Exception e) { return new Result {Message="Falha local/rede: "+e.GetType().Name+"; caminho local com "+target.Length+" caracteres. Arquivos existentes preservados."}; }
            return new Result {Message="Nenhuma rota disponivel."};
        }
        Result One(string url,string target,long size,string identity,CookieContainer cookies) {
            target=LocalPaths.Native(target);
            Directory.CreateDirectory(Path.GetDirectoryName(target));
            string part=target+".spdownload.part", meta=part+".json";
            Partial saved=null;
            if(File.Exists(meta)) { try { saved=Read(meta); } catch { } }
            if(File.Exists(part) && saved==null) return new Result {Message="Parcial sem identificacao; preserve ou mova antes de repetir."};
            long offset=File.Exists(part)?new FileInfo(part).Length:0;
            if(saved==null || saved.Identity!=identity || saved.Size!=size || offset>=size || String.IsNullOrEmpty(saved.ETag) || saved.ETag.StartsWith("W/")) offset=0;
            var req=(HttpWebRequest)WebRequest.Create(url);
            req.CookieContainer=cookies; req.AllowAutoRedirect=false;
            req.Timeout=20000; req.ReadWriteTimeout=60000; req.Headers["Accept-Encoding"]="identity";
            if(offset>0) { req.AddRange(offset); req.Headers["If-Range"]=saved.ETag; Events.Enqueue("PARCIAL_RETOMADA_SOLICITADA offset="+offset); }
            HttpWebResponse response;
            try { response=(HttpWebResponse)req.GetResponse(); }
            catch(WebException e) {
                if(e.Response!=null) using(var error=(HttpWebResponse)e.Response) return HttpError(error);
                return new Result {Message="Falha de rede: "+e.Status};
            }
            using(response) {
                int status=(int)response.StatusCode;
                if(status!=200 && status!=206) return HttpError(response);
                if((response.ContentType??"").IndexOf("text/html",StringComparison.OrdinalIgnoreCase)>=0 &&
                    !target.EndsWith(".html",StringComparison.OrdinalIgnoreCase) && !target.EndsWith(".htm",StringComparison.OrdinalIgnoreCase))
                    return new Result {Status=status,Authentication=true,Message="HTML inesperado; confira o login."};
                if(status==206) {
                    var m=Regex.Match(response.Headers["Content-Range"]??"",@"^bytes (\d+)-(\d+)/(\d+)$");
                    if(offset==0 || !m.Success || long.Parse(m.Groups[1].Value)!=offset || long.Parse(m.Groups[3].Value)!=size ||
                       long.Parse(m.Groups[2].Value)<offset || long.Parse(m.Groups[2].Value)>=size || response.Headers["ETag"]!=saved.ETag)
                        return new Result {Message="Content-Range/ETag invalido; parcial preservado."};
                    Events.Enqueue("PARCIAL_RETOMADA_CONFIRMADA offset="+offset);
                } else { if(offset>0) Events.Enqueue("PARCIAL_REINICIO servidor_retornou_200"); offset=0; }
                Save(meta,new Partial {Identity=identity,Size=size,ETag=response.Headers["ETag"]});
                long written=offset;
                Interlocked.Exchange(ref Bytes,written);
                using(var input=response.GetResponseStream())
                using(var output=new FileStream(part,offset>0?FileMode.Append:FileMode.Create,FileAccess.Write,FileShare.None)) {
                    byte[] buffer=new byte[1048576]; int count;
                    while((count=input.Read(buffer,0,buffer.Length))>0) {
                        if(Cancel) throw new OperationCanceledException();
                        output.Write(buffer,0,count); written+=count; Interlocked.Exchange(ref Bytes,written);
                        if(written>size) break;
                    }
                }
                if(Cancel) throw new OperationCanceledException();
                if(written!=size) {
                    Save(meta,new Partial {Identity=identity,Size=size,ETag=null});
                    return new Result {Status=status,SizeMismatch=true,Expected=size,Received=written,ResponseLength=response.ContentLength,
                        ResponseETag=response.Headers["ETag"],ContentEncoding=response.ContentEncoding,
                        Message="Tamanho divergente: esperado "+size+" bytes, gravado "+written+" bytes; parcial preservado para reinicio."};
                }
                if(File.Exists(target) || Directory.Exists(target)) return new Result {Message="Destino criado durante transferencia; preservado."};
                File.Move(part,target); File.Delete(meta);
                return new Result {Success=true,Status=status,Message="Validado por tamanho."};
            }
        }
        public static string Probe(string url,CookieContainer cookies) {
            try {
                var req=(HttpWebRequest)WebRequest.Create(url); req.CookieContainer=cookies;
                req.AllowAutoRedirect=false; req.Timeout=20000; req.ReadWriteTimeout=20000;
                req.Headers["Accept-Encoding"]="identity";
                HttpWebResponse response;
                try { response=(HttpWebResponse)req.GetResponse(); }
                catch(WebException e) { if(e.Response==null) return "rede="+e.Status; response=(HttpWebResponse)e.Response; }
                using(response) {
                    int n=0,k; byte[] buffer=new byte[32768];
                    using(var stream=response.GetResponseStream()) while(n<buffer.Length && (k=stream.Read(buffer,n,buffer.Length-n))>0) n+=k;
                    return "HTTP="+(int)response.StatusCode+" tipo="+Clean(response.ContentType)+" amostra="+n+" zip="+(n>=4 && buffer[0]==80 && buffer[1]==75 && buffer[2]==3 && buffer[3]==4);
                }
            } catch(Exception e) { return "falha="+e.GetType().Name; }
        }
    }
}
