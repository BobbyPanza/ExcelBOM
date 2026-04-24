namespace ExcelBOM;

public class Anagrafica
{
    public string PACOD { get; set; } = "";
    public string PADSC { get; set; } = "";
    public string FMCOD { get; set; } = "";
    public string PAUDM { get; set; } = "";
    public string PATYP { get; set; } = "";
    public string PATYF { get; set; } = "";
}

public class Distinta
{
    public string DBPAR { get; set; } = "";
    public string DBCHL { get; set; } = "";
    public string DBSEQ { get; set; } = "";
    public string DBQTA { get; set; } = "";
    public string DBUDM { get; set; } = "";
    public string DBNT1 { get; set; } = "";
}

public class Fase
{
    public string PACOD { get; set; } = "";
    public string FASEQ { get; set; } = "";
    public string FACOD { get; set; } = "";
    public string FADSC { get; set; } = "";
    public string FAHLV { get; set; } = "";
    public string FASLV { get; set; } = "";
}