#include "pch.h"

using namespace std;
using namespace Spire::Pdf;
using namespace Spire::Common;

intrusive_ptr<PdfPageTemplateElement> CreateHeaderTemplate(intrusive_ptr<PdfDocument> doc, intrusive_ptr<PdfMargins> margins, intrusive_ptr<SizeF> pageSize)
{
	//create a PdfPageTemplateElement object as header space
	intrusive_ptr<PdfPageTemplateElement> headerSpace = new PdfPageTemplateElement(pageSize->GetWidth(), margins->GetTop());
	headerSpace->SetForeground(false);

	//declare two float variables
	float x = margins->GetLeft();
	float y = 0;

	wstring inputFile = DATAPATH"E-iceblueLogo.png";
	//draw image in header space 
	intrusive_ptr<PdfImage> headerImage = PdfImage::FromFile(inputFile.c_str());
	float width = headerImage->GetWidth() / 2;
	float height = headerImage->GetHeight() / 2;
	headerSpace->GetGraphics()->DrawImage(headerImage, x, margins->GetTop() - height - 5, width, height);

	//draw line in header space
	intrusive_ptr<PdfPen> pen = new PdfPen(PdfBrushes::GetLightGray(), 1);
	headerSpace->GetGraphics()->DrawLine(pen, x, y + margins->GetTop() - 2, pageSize->GetWidth() - x, y + margins->GetTop() - 2);

	return headerSpace;
}

intrusive_ptr<PdfPageTemplateElement> CreateFooterTemplate(intrusive_ptr<PdfDocument> doc, intrusive_ptr<PdfMargins> margins, intrusive_ptr<SizeF> pageSize)
{
	//create a PdfPageTemplateElement object which works as footer space
	intrusive_ptr<PdfPageTemplateElement> footerSpace = new PdfPageTemplateElement(pageSize->GetWidth(), margins->GetBottom());
	footerSpace->SetForeground(false);

	//declare two float variables
	float x = margins->GetLeft();
	float y = 0;

	//draw line in footer space
	intrusive_ptr<PdfPen> pen = new PdfPen(PdfBrushes::GetGray(), 1);
	footerSpace->GetGraphics()->DrawLine(pen, x, y, pageSize->GetWidth() - x, y);

	//draw text in footer space
	y = y + 5;

	LPCWSTR_S arial = L"Arial";
	intrusive_ptr<PdfTrueTypeFont> font = new PdfTrueTypeFont(arial, 10.f, PdfFontStyle::Regular, true);

	//draw dynamic field in footer space
	intrusive_ptr<PdfPageNumberField> number = new PdfPageNumberField();
	intrusive_ptr<PdfPageCountField> count = new PdfPageCountField();
	vector< intrusive_ptr<PdfAutomaticField>> list = { number,count };
	intrusive_ptr<PdfCompositeField> compositeField = new PdfCompositeField(font, PdfBrushes::GetBlack(), L"Page {0} of {1}", list);
	compositeField->SetStringFormat(new PdfStringFormat(PdfTextAlignment::Left, PdfVerticalAlignment::Top));
	intrusive_ptr<SizeF> size = font->MeasureString(compositeField->GetText());
	compositeField->SetBounds(new RectangleF(x, y, size->GetWidth(), size->GetHeight()));
	Object::Convert<PdfGraphicsWidget>(compositeField)->Draw(footerSpace->GetGraphics());

	return footerSpace;
}

int main()
{
	wstring outputFile = OUTPUTPATH"ImageAndPageNumber.pdf";

	//create a PDF document
	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->GetPageSettings()->SetSize(PdfPageSize::A4());

	//reset the default margins to 0
	doc->GetPageSettings()->SetMargins(new PdfMargins(0));

	//create a PdfMargins object, the parameters indicate the page margins you want to set
	intrusive_ptr<PdfMargins> margins = new PdfMargins(50, 50, 50, 50);

	//get page size
	intrusive_ptr<SizeF> pageSize = doc->GetPageSettings()->GetSize();

	//create a header template with content and apply it to page template
	doc->GetTemplate()->SetTop(CreateHeaderTemplate(doc, margins, pageSize));

	//create a footer template with content and apply it to page template
	doc->GetTemplate()->SetBottom(CreateFooterTemplate(doc, margins, pageSize));

	//apply blank templates to other parts of page template
	doc->GetTemplate()->SetLeft(new PdfPageTemplateElement(margins->GetLeft(), doc->GetPageSettings()->GetSize()->GetHeight()));
	doc->GetTemplate()->SetRight(new PdfPageTemplateElement(margins->GetRight(), doc->GetPageSettings()->GetSize()->GetHeight()));
	//Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}