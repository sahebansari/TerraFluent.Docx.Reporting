namespace TerraFluent.Docx.Reporting.Infra;

public interface IColumnDescriptor
{
    IColumnDescriptor Spacing(float points);
    IContainer Item();
}
