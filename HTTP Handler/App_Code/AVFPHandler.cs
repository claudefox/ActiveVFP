using System;
using System.Web;
using System.Web.UI;
using System.Web.SessionState;
using activevfp_dotnetproxy;
public class AVFPHandler : Page, IHttpAsyncHandler, IRequiresSessionState
{



    public IAsyncResult BeginProcessRequest(HttpContext context, AsyncCallback callback, object extraData)
    {

        return AspCompatBeginProcessRequest(context, callback, extraData);

    }



    public void EndProcessRequest(IAsyncResult result)
    {

        AspCompatEndProcessRequest(result);

    }



    /// <summary>

    /// Overrides Page.OnInit which is called as part of Async callbacks

    /// </summary>

    /// <param name="e"></param>     

    protected override void OnInit(EventArgs e) 
   
{

        /// Call my base handler
        this.ProcessRequest(HttpContext.Current);

        HttpContext.Current.ApplicationInstance.CompleteRequest();

    }

    //protected override void OnInit(EventArgs e)
    
    public void ProcessRequest(HttpContext context)
    {
        HttpRequest Request = context.Request;
        HttpResponse Response = context.Response;
     
        // This handler is called whenever a file ending 
        // in .avfp is requested. A file with that extension
        // does not need to exist.
        
        server x;
        x = new server();
        try
        {
            Response.Write(x.Process());
        }
        catch (Exception ex)
        {
            Response.Write("Caught .NET exception, source: " + ex.Source + " message: " + ex.Message);
        }

    }
    public bool IsReusable
    {
        // To enable pooling, return true here.
        // This keeps the handler in memory.
        get { return true; }
    }
   

}


