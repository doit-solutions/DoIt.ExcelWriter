namespace DoIt.ExcelWriter;

[AttributeUsage(AttributeTargets.Property)]
public class ExcelColumnAttribute : Attribute
{
    public string? Title { get; set; }
    public bool Ignore { get; set; }
    public double CustomWidth { get; set; }

    public ExcelColumnAttribute()
    {
    }

    public ExcelColumnAttribute(string title)
    {
        Title = title;
    }
}
