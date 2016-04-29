'EXAMPLE OF PULLING BLOOMBERG DATA IN VBA
Option Explicit

Dim curRow As Integer
Public bbdata As Object

Public Function GetBLPLink(ticker As String, RetType As Integer, bb_type As String)

    Dim DefaultVal As String    'Value when blp comes out bad
    Dim bbergCode As String

    ticker = CheckTickerException(ticker)
    
    Select Case RetType
        Case 1:
            bbergCode = "EQY_SH_OUT_TOT_MULT_SH"
        Case 2:
            bbergCode = "LAST_TRADE"
        Case 3:
            bbergCode = "PX_LAST"
        Case 4:
            bbergCode = "EQY_RAW_BETA_6M"
            DefaultVal = "1"
        Case 5:
            bbergCode = "EXPECTED_REPORT_DT"
            DefaultVal = ""
        Case 6:
            bbergCode = "EXPECTED_REPORT_TIME"
            DefaultVal = ""
        Case 7:
            bbergCode = "DVD_SH_LAST" '"IS_DIV_PER_SHR"
            DefaultVal = ""
        Case 8:
            bbergCode = "DVD_PAY_DT"
            DefaultVal = ""
        Case 9:     '52 Week high
            bbergCode = "HIGH_52WEEK"
        Case 10:    '52 week Low
            bbergCode = "LOW_52WEEK"
        Case 11:
            bbergCode = "DVD_EX_DT"
            DefaultVal = ""
        Case 12:
            bbergCode = "DVD_RECORD_DT"
            DefaultVal = ""
        Case 13:
            bbergCode = "EQY_FLOAT"
            DefaultVal = "0"
        Case 14:
            bbergCode = "SHORT_INT"
            DefaultVal = "0"
        Case 15:
            bbergCode = "VOLUME_AVG_20D"
            DefaultVal = "0"
        Case 16:
            bbergCode = "HIGH_DT_52WEEK"
        Case 17:
            bbergCode = "VOLUME"
        Case 18:
            bbergCode = "BEST_EEPS_LAST_QTR"
        Case 19:
            bbergCode = "BEST_AEPS_LST_QTR"
        Case 20:
            bbergCode = "OPT_DELTA"
        Case 21:
            bbergCode = "ALL_DAY_VWAP"
        Case 22:
            bbergCode = "PX_CLOSE_1D"
        Case 23:
            bbergCode = "PX_YEST_CLOSE"
        Case 24:
            bbergCode = "PX_HIGH"
        Case 25:
            bbergCode = "PX_LOW"
        Case 26:
            bbergCode = "EQY_FREE_FLOAT_PCT"
        Case 27:
            bbergCode = "5D_AVERAGE_CURRENT_VOLUME"
        Case 28:    'Market Cap
            bbergCode = "CUR_MKT_CAP"
        Case 29:
            bbergCode = "CHG_PCT_5D"
        Case 30:
            bbergCode = "CHG_PCT_1M"
        Case 31:
            bbergCode = "CHG_PCT_YTD"
        Case 32:
            bbergCode = "OPT_UNDL_TICKER"
        Case 33:
            bbergCode = "VWAP"
        Case Else
            bbergCode = ""
            'BEST_AEPS_LST_QTR - BEST_EEPS_LST_QTR
    End Select
        
    If RetType <> 33 Then
        GetBLPLink = "=BDP(" & """" & ticker & " Equity" & """" & "," & """" & bbergCode & """" & ")"
    Else
        GetBLPLink = "=BDP(" & """" & ticker & " Equity" & """" & "," & """" & bbergCode & """" & "," & """" & "3:45 PM" & """" & "," & """" & "4:00 PM" & """" & ")"
    End If
End Function

Public Function GetBlpData3(ticker As String, RetType As Integer, bb_type As String, Optional override1 As String) As String
    Dim res As Variant
    Dim bbergCode As String
    Dim DefaultVal As String    'Value when blp comes out bad
    
    On Error GoTo ErrHandler
    ticker = CheckTickerException(ticker)
        
    If bbdata Is Nothing Then
      Set bbdata = New clsBlpControl2
    End If
    
    Select Case RetType
        Case 1:
            bbergCode = "EQY_SH_OUT_TOT_MULT_SH"
        Case 2:
            bbergCode = "PX_LAST"
        Case 3:
            bbergCode = "RTG_SP"
        Case 4:
            bbergCode = "EQY_RAW_BETA_6M"
            DefaultVal = "1"
        Case 5:
            bbergCode = "EXPECTED_REPORT_DT"
            DefaultVal = ""
        Case 6:
            bbergCode = "EXPECTED_REPORT_TIME"
            DefaultVal = ""
        Case 7:
            bbergCode = "DVD_SH_LAST" '"IS_DIV_PER_SHR"
            DefaultVal = ""
        Case 8:
            bbergCode = "DVD_PAY_DT"
            DefaultVal = ""
        Case 9:     '52 Week high
            bbergCode = "HIGH_52WEEK"
        Case 10:    '52 week Low
            bbergCode = "LOW_52WEEK"
        Case 11:
            bbergCode = "DVD_EX_DT"
            DefaultVal = ""
        Case 12:
            bbergCode = "DVD_RECORD_DT"
            DefaultVal = ""
        Case 13:
            bbergCode = "EQY_FLOAT"
            DefaultVal = "0"
        Case 14:
            bbergCode = "SHORT_INT"
            DefaultVal = "0"
        Case 15:
            bbergCode = "VOLUME_AVG_20D"
            DefaultVal = "0"
        Case 16:
            bbergCode = "HIGH_DT_52WEEK"
        Case 17:
            bbergCode = "EPS_SURPRISE_LAST_QTR"
        Case 18:
            bbergCode = "BEST_EEPS_LAST_QTR"
        Case 19:
            bbergCode = "BEST_AEPS_LST_QTR"
        Case 20:
            bbergCode = "OPT_DELTA"
        Case 21:
            bbergCode = "ALL_DAY_VWAP"
        Case 22:
            bbergCode = "PX_CLOSE"
        Case 23:
            bbergCode = "PX_YEST_CLOSE"
        Case 24:
            bbergCode = "PX_HIGH"
        Case 25:
            bbergCode = "PX_LOW"
        Case 26:
            bbergCode = "EQY_FREE_FLOAT_PCT"
        Case 27:
            bbergCode = "BEST_EEPS_CUR_YR"
        Case 28:
            bbergCode = "BEST_EPS_NXT_YR"
        Case 29:
            bbergCode = "BEST_TARGET_PRICE"
        Case 30:
            bbergCode = "EQY_SH_OUT"
            DefaultVal = "0"
        Case 31:
            bbergCode = "CHG_PCT_1D"
        Case 32:
            bbergCode = "PX_CLOSE_2D"
        Case 33:
            bbergCode = "PX_YEST_HIGH"
        Case 34:
            bbergCode = "PX_YEST_LOW"
        Case 35:
            bbergCode = "OPT_UNDL_TICKER"
        Case 36:
            bbergCode = "OPT_DELTA_MID"
        Case 37:
            bbergCode = "CHG_PCT_2D"
        Case 38:
            bbergCode = "VOLATILITY_260D"
        Case 39:
            bbergCode = "GICS_SECTOR_NAME"
        Case 40:
        
          bbergCode = "EQY_BETA_RAW_OVERRIDABLE"
          'Work here for override in new format...
        
          Dim edate As String
          Dim sdate As String
          Dim prev_day As String
            
          prev_day = GetPrevBizDay(Date)
          edate = Format(prev_day, "YYYYMMDD")
          sdate = Format(DateAdd("m", -6, prev_day), "YYYYMMDD")
            
        Case 41:
            bbergCode = "CNTRY_ISSUE_ISO"
        Case 42:
            bbergCode = "CUR_MKT_CAP"
        Case 43:
            bbergCode = "TICKER"
        Case 44:
            bbergCode = "GICS_INDUSTRY_GROUP_NAME"
        Case 45:
            bbergCode = "GICS_INDUSTRY_NAME"        '"GICS_SUB_INDUSTRY_NAME",
        Case Else
            bbergCode = ""
    End Select

    
        bbdata.BlpSubscribe
        bbdata.req.GetElement("securities").AppendValue ticker + " " + bb_type
        bbdata.req.GetElement("fields").AppendValue bbergCode
    
        If override1 <> "" Then
            Dim overrides As Element
            Dim ov1 As Element
            Dim ov2 As Element
            Dim ov3 As Element
            Dim ov4 As Element

            Set overrides = bbdata.req.GetElement("overrides")
        
            Set ov1 = overrides.AppendElment()
            ov1.SetElement "fieldId", "EQY_BETA_OVERRIDE_REL_INDEX"
            ov1.SetElement "value", override1
            
            Set ov2 = overrides.AppendElment()
            Call ov2.SetElement("fieldId", "EQY_BETA_OVERRIDE_PERIOD")
            Call ov2.SetElement("value", "D")
        
            Set ov3 = overrides.AppendElment()
            Call ov3.SetElement("fieldId", "EQY_BETA_OVERRIDE_START_DT")
            Call ov3.SetElement("value", sdate)
        
            Set ov4 = overrides.AppendElment()
            Call ov4.SetElement("fieldId", "EQY_BETA_OVERRIDE_END_DT")
            Call ov4.SetElement("value", edate)
        End If
        
        res = bbdata.BlpSendRequest
        
        If Not IsEmpty(res(0, 0)) Then
            GetBlpData3 = Format(res(0, 0))
        Else
            GetBlpData3 = DefaultVal
        End If
    
    Exit Function
ErrHandler:
    GetBlpData3 = DefaultVal
End Function
