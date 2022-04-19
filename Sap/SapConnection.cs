using System.Runtime.InteropServices;
using System.Reflection;
using Microsoft.VisualBasic;

namespace Sap;

/// <summary>Gerencia a conexão com SAP GUI</summary>
public class SapConnection
{
    public object? SapGui { get; set; }
    public object? SapApp { get; set; }
    public object? SapCon { get; set; }

    /// <summary>Armazena a sessão SAP GUI atual</summary>
    public object? SapCurrentSession { get; set; }

    /// <summary>Armazena todas as sessões abertas no SAP Gui</summary>
    public List<object>? SapSessions { get; set; }

    public SapConnection()
    {
        this.SapGui = null;
        this.SapApp = null;
        this.SapCon = null;
        this.SapCurrentSession = null;
        this.SapSessions = new List<object>();
    }

    /// <summary>Executa o(s) método(s)</summary>
    public dynamic? InvokeMethod(object? target, string methodName, object[]? methodParams = null)
    {
        return target?.GetType()
            .InvokeMember(methodName, BindingFlags.InvokeMethod, null, target, methodParams);
    }

    /// <summary>Obtém a(s) propriedade(s)</summary>
    public dynamic? GetProperty(object? target, string propertyName, object[]? propertyParams = null)
    {
        return target?.GetType()
            .InvokeMember(propertyName, BindingFlags.GetProperty, null, target, propertyParams);
    }
    /// <summary>Defini a(s) propriedade(s)</summary>
    public dynamic? SetProperty(object? target, string propertyName, object[]? propertyParams = null)
    {
        return target?.GetType()
            .InvokeMember(propertyName, BindingFlags.SetProperty, null, target, propertyParams);
    }

    /// <summary>Cria a conexão com o SAP GUI</summary>
    public void Connection()
    {
        try
        {
            this.SapGui = Interaction.GetObject("SAPGUI");
            this.SapApp = this.InvokeMethod(this.SapGui, "GetScriptingEngine");
            this.SapCon = this.GetProperty(this.SapApp, "Children", new object[]{0});
        }
        catch (Exception e)
        {
            this.Close();
            throw new Exception("[CONNECTION FAILED]: " + e.Message);
        }
    }

    /// <summary>Obtém a sessão SAP GUI atual</summary>
    public void GetCurrentSession()
    {
        try
        {
            if (this.SapCon != null)
            {
                this.SapCurrentSession = this.GetProperty(this.SapCon, "Children", new object[]{0});
            }
        }
        catch (Exception e)
        {
            this.Close();
            throw new Exception("[GET CURRENT SESSION FAILED]: " + e.Message);
        }
    }

    /// <summary>Obtém todas as sessões SAP GUI abertas</summary>
    public void GetAllSessions()
    {
        try
        {
            if (this.SapCon != null)
            {
                if (this.SapCon != null)
                {
                    var Sessions = this.GetProperty(this.SapCon, "Sessions");
                    if (Sessions != null)
                    {
                        foreach (var session in Sessions)
                        {
                            if (session != null)
                                this.SapSessions?.Add(session);
                        }
                    }
                }
            }
        }
        catch (Exception e)
        {
            this.Close();
            throw new Exception("[GET ALL SESSIONS FAILED]: " + e.Message);
        }
    }

    /// <summary>Obtém a transação da sessão passada como parâmetro</summary>
    public string? GetTransaction(object? Session)
    {
        string? transaction = null;
        try
        {
            if (Session != null)
            {
                dynamic? info = this.GetProperty(this.SapCurrentSession, "Info");
                transaction = this.GetProperty(info, "Transaction");
            }
        }
        catch (Exception e)
        {
            this.Close();
            throw new Exception("[GET TRANSACTION FAILED]: " + e.Message);
        }
        return transaction;
    }

    ///<summary>Libera os objetos COM</summary>
    public void Close()
    {
        if (this.SapSessions != null)
        {
            this.SapSessions.ForEach(session => {
                if (session != null)
                {
                    Marshal.ReleaseComObject(session);
                }
            });
            this.SapSessions.Clear();
        }
        if (this.SapCurrentSession != null)
        {
            Marshal.ReleaseComObject(this.SapCurrentSession);
            this.SapCurrentSession = null;
        }
        if (this.SapCon != null)
        {
            Marshal.ReleaseComObject(this.SapCon);
            this.SapCon = null;
        }
        if (this.SapApp != null)
        { 
            Marshal.ReleaseComObject(this.SapApp);
            this.SapApp = null;
        }
        if (this.SapGui != null)
        {
            Marshal.ReleaseComObject(this.SapGui);
            this.SapGui = null;
        }
    }
}
