namespace TerraFluent.Docx.Reporting.Infra;

public interface IRowDescriptor
{
    IRowDescriptor Spacing(float points);
    IContainer RelativeItem(float size = 1);
    IContainer AutoItem();
    IContainer ConstantItem(float widthPoints);
}
