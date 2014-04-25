using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

public struct S_MethodResult
{
    public bool Success;
    public int ReturnCode;
    public string Message;
    public object ReturnValue;
    public S_MethodResult(bool pSuccess, int pReturnCode, string pMessage, object pReturnValue)
    {
        this.Success = pSuccess;
        this.ReturnCode = pReturnCode;
        this.Message = pMessage;
        this.ReturnValue = pReturnValue;
    }

    public void SetMethodResult(bool pSuccess)
    {
        this.Success = pSuccess;
        if (pSuccess)
        {
            this.ReturnCode = 0;
        }
        else
        {
            this.ReturnCode = -1;
        }
        this.Message = string.Empty;
        this.ReturnValue = null;
    }

    public void SetMethodResult(bool pSuccess, int pReturnCode)
    {
        this.Success = pSuccess;
        this.ReturnCode = pReturnCode;
        this.Message = string.Empty;
        this.ReturnValue = null;
    }

    public void SetMethodResult(bool pSuccess, int pReturnCode, string pMessage)
    {
        this.Success = pSuccess;
        this.ReturnCode = pReturnCode;
        this.Message = pMessage;
        this.ReturnValue = null;
    }

    public void SetMethodResult(bool pSuccess, int pReturnCode, string pMessage, object pReturnValue)
    {
        this.Success = pSuccess;
        this.ReturnCode = pReturnCode;
        this.Message = pMessage;
        this.ReturnValue = pReturnValue;
    }
}