namespace Nedev.DocToDocx.Utils;

/// <summary>
/// Helper for parsing Word DTTM structures (32-bit values representing dates/times).
/// MS-DOC §2.9.89: DTTM (4 bytes)
/// Bit 0-5: Minutes (0-59)
/// Bit 6-10: Hours (0-23)
/// Bit 11-15: Day of month (1-31)
/// Bit 16-19: Month (1-12)
/// Bit 20-28: Year (0-511, relative to 1900)
/// Bit 29-31: Day of week (0-6, nullable)
/// </summary>
public static class DttmHelper
{
    public static DateTime ParseDttm(uint dttm)
    {
        if (dttm == 0) return DateTime.Now;
        try 
        {
            int mint = (int)(dttm & 0x3F);
            int hr = (int)((dttm >> 6) & 0x1F);
            int dom = (int)((dttm >> 11) & 0x1F);
            int mon = (int)((dttm >> 16) & 0x0F);
            int yr = 1900 + (int)((dttm >> 20) & 0x1FF);
            
            // Clamp values for safety before DateTime constructor
            int validYr = Math.Clamp(yr, 1900, 2100);
            int validMon = Math.Max(1, Math.Min(12, mon));
            int validDom = Math.Max(1, Math.Min(DateTime.DaysInMonth(validYr, validMon), dom));
            int validHr = Math.Clamp(hr, 0, 23);
            int validMin = Math.Clamp(mint, 0, 59);

            return new DateTime(validYr, validMon, validDom, validHr, validMin, 0);
        } 
        catch 
        {
            return DateTime.Now;
        }
    }
}
