#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

inline PdfFontStyle operator|(PdfFontStyle a, PdfFontStyle b) {
	return static_cast<PdfFontStyle>(static_cast<int>(a) | static_cast<int>(b));
}
	
	float AddDocumentLinkAnnotation(intrusive_ptr<PdfPageBase> page, float y)
	{
		intrusive_ptr<PdfTrueTypeFont> font = new PdfTrueTypeFont(L"Arial", 12.f, PdfFontStyle::Regular, true);
		intrusive_ptr<PdfStringFormat> format = new PdfStringFormat();
		format->SetMeasureTrailingSpaces(true);
		wstring prompt = L"Document Link: ";
		intrusive_ptr<SizeF> size = font->MeasureString(prompt.c_str());
		page->GetCanvas()->DrawString(prompt.c_str(), font, PdfBrushes::GetDodgerBlue(), 0, y);
		float x = font->MeasureString(prompt.c_str(), format)->GetWidth();
		intrusive_ptr<PdfDestination> dest = new PdfDestination(page);
		dest->SetMode(PdfDestinationMode::Location);
		dest->SetLocation(new PointF(0, y));
		dest->SetZoom(2.f);

		wstring label = L"Click me, Zoom 200%";
		size = font->MeasureString(label.c_str());
		intrusive_ptr<RectangleF> bounds = new RectangleF(x, y, size->GetWidth(), size->GetHeight());
		page->GetCanvas()->DrawString(label.c_str(), font, PdfBrushes::GetOrangeRed(), x, y);
		intrusive_ptr<PdfDocumentLinkAnnotation> annotation = new PdfDocumentLinkAnnotation(bounds, dest);
		annotation->SetColor(new PdfRGBColor(Spire::Common::Color::GetBlue()));
		Object::Convert<PdfNewPage>(page)->GetAnnotations()->Add(annotation);
		y = bounds->GetBottom();
		return y;
	}

	float AddFileLinkAnnotation(intrusive_ptr<PdfPageBase> page, float y)
	{

		intrusive_ptr<PdfTrueTypeFont> font = new PdfTrueTypeFont(L"Arial", 12.f, PdfFontStyle::Regular, true);
		intrusive_ptr<PdfStringFormat> format = new PdfStringFormat();
		format->SetMeasureTrailingSpaces(true);
		wstring prompt = L"Launch File: ";
		intrusive_ptr<SizeF> size = font->MeasureString(prompt.c_str());
		page->GetCanvas()->DrawString(prompt.c_str(), font, PdfBrushes::GetDodgerBlue(), 0, y);
		float x = font->MeasureString(prompt.c_str(), format)->GetWidth();

		wstring label = L"Launch Notepad.exe";
		size = font->MeasureString(label.c_str());
		intrusive_ptr<RectangleF> bounds = new RectangleF(x, y, size->GetWidth(), size->GetHeight());
		page->GetCanvas()->DrawString(label.c_str(), font, PdfBrushes::GetOrangeRed(), x, y);
		intrusive_ptr<PdfFileLinkAnnotation> annotation = new PdfFileLinkAnnotation(bounds, L"C:\\Windows\\Notepad.exe");
		annotation->SetColor(new PdfRGBColor(Spire::Common::Color::GetBlue()));
		Object::Convert<PdfNewPage>(page)->GetAnnotations()->Add(annotation);
		y = bounds->GetBottom();
		return y;
	}

	float AddFreeTextAnnotation(intrusive_ptr<PdfPageBase> page, float y)
	{
		intrusive_ptr<PdfTrueTypeFont> font = new PdfTrueTypeFont(L"Arial", 12.f, PdfFontStyle::Regular, true);
		intrusive_ptr<PdfStringFormat> format = new PdfStringFormat();
		format->SetMeasureTrailingSpaces(true);
		wstring prompt = L"Text Markup: ";
		intrusive_ptr<SizeF> size = font->MeasureString(prompt.c_str());
		page->GetCanvas()->DrawString(prompt.c_str(), font, PdfBrushes::GetDodgerBlue(), 0, y);
		float x = font->MeasureString(prompt.c_str(), format)->GetWidth();

		wstring label = L"I'm a text box, not a TV";
		size = font->MeasureString(label.c_str());
		intrusive_ptr<RectangleF> bounds = new RectangleF(x, y, size->GetWidth(), size->GetHeight());
		page->GetCanvas()->DrawRectangle(new PdfPen(new PdfRGBColor(Spire::Common::Color::GetBlue()), 0.1f), bounds);
		page->GetCanvas()->DrawString(label.c_str(), font, PdfBrushes::GetOrangeRed(), x, y);
		intrusive_ptr<PointF> location = new PointF(bounds->GetRight() + 16, bounds->GetTop() - 16);
		intrusive_ptr<RectangleF> annotaionBounds = new RectangleF(location, new SizeF(80, 32));
		intrusive_ptr<PdfFreeTextAnnotation> annotation = new PdfFreeTextAnnotation(annotaionBounds);
		annotation->SetAnnotationIntent(PdfAnnotationIntent::FreeTextCallout);
		annotation->SetBorder(new PdfAnnotationBorder(0.5f));
		annotation->SetBorderColor(new PdfRGBColor(Spire::Common::Color::GetRed()));
		location = new PointF(bounds->GetRight() + 16, bounds->GetTop() - 16);
		intrusive_ptr<PointF> pointF[] = {
					new PointF(bounds->GetRight(), bounds->GetTop()),
					new PointF(bounds->GetRight() + 8, bounds->GetTop() - 8),
					location };
		vector<intrusive_ptr<PointF>> vec;
		vec.insert(vec.begin(), pointF, pointF + 3);
		annotation->SetCalloutLines(vec);
		annotation->SetColor(new PdfRGBColor(Spire::Common::Color::GetYellow()));
		annotation->SetFlags(PdfAnnotationFlags::Locked);
		annotation->SetFont(font);
		annotation->SetLineEndingStyle(PdfLineEndingStyle::OpenArrow);
		annotation->SetMarkupText(L"Just a joke.");
		annotation->SetOpacity(0.75f);
		annotation->SetTextMarkupColor(new PdfRGBColor(Spire::Common::Color::GetGreen()));
		Object::Convert<PdfNewPage>(page)->GetAnnotations()->Add(annotation);
		y = bounds->GetBottom();
		return y;
	}

	float AddLineAnnotation(intrusive_ptr<PdfPageBase> page, float y)
	{
		intrusive_ptr<PdfTrueTypeFont> font = new PdfTrueTypeFont(L"Arial", 12.f, PdfFontStyle::Regular, true);
		intrusive_ptr<PdfStringFormat> format = new PdfStringFormat();
		format->SetMeasureTrailingSpaces(true);
		wstring prompt = L"Line Annotation: ";
		intrusive_ptr<SizeF> size = font->MeasureString(prompt.c_str());
		page->GetCanvas()->DrawString(prompt.c_str(), font, PdfBrushes::GetDodgerBlue(), 0, y);
		float x = font->MeasureString(prompt.c_str(), format)->GetWidth();
		wstring label = L"Line Anotation";
		size = font->MeasureString(label.c_str());
		page->GetCanvas()->DrawString(label.c_str(), font, PdfBrushes::GetOrangeRed(), x, y);
		intrusive_ptr<RectangleF> bounds = new RectangleF(x, y, size->GetWidth(), size->GetHeight());
		int* linePoints = new int[]
		{
			(int)bounds->GetLeft(), (int)bounds->GetTop(), (int)bounds->GetRight(), (int)bounds->GetBottom()
		};
		vector<int> vec;
		vec.insert(vec.begin(), linePoints, linePoints + 4);
		intrusive_ptr<PdfLineAnnotation> annotation = new PdfLineAnnotation(vec, L"Annotation");
		annotation->SetBeginLineStyle(PdfLineEndingStyle::ClosedArrow);
		annotation->SetEndLineStyle(PdfLineEndingStyle::ClosedArrow);
		annotation->SetLineCaption(true);
		annotation->SetBackColor(new PdfRGBColor(Spire::Common::Color::GetBlack()));
		annotation->SetCaptionType(PdfLineCaptionType::Inline);
		Object::Convert<PdfNewPage>(page)->GetAnnotations()->Add(annotation);
		y = bounds->GetBottom();
		return y;
	}

	float AddTextMarkupAnnotation(intrusive_ptr<PdfPageBase> page, float y)
	{
		intrusive_ptr<PdfTrueTypeFont> font = new PdfTrueTypeFont(L"Arial", 12.f, PdfFontStyle::Regular, true);
		intrusive_ptr<PdfStringFormat> format = new PdfStringFormat();
		format->SetMeasureTrailingSpaces(true);
		wstring prompt = L"Highlight incorrect spelling: ";
		intrusive_ptr<SizeF> size = font->MeasureString(prompt.c_str(), format);
		page->GetCanvas()->DrawString(prompt.c_str(), font, PdfBrushes::GetDodgerBlue(), 0, y);
		float x = size->GetWidth();
		wstring label = L"demo of anotation";
		page->GetCanvas()->DrawString(label.c_str(), font, PdfBrushes::GetOrangeRed(), x, y);
		size = font->MeasureString(L"demo of ", format);
		x += size->GetWidth();
		intrusive_ptr<PointF> incorrectWordLocation = new PointF(x, y);
		wstring markupText = L"Should be 'annotation'";
		intrusive_ptr<PdfTextMarkupAnnotation> annotation
			= new PdfTextMarkupAnnotation(markupText.c_str(), L"anotation", new RectangleF(x, y, 100.f, 100.f), font);
		annotation->SetTextMarkupAnnotationType(PdfTextMarkupAnnotationType::Highlight);
		annotation->SetTextMarkupColor(new PdfRGBColor(Spire::Common::Color::GetLightSkyBlue()));
		Object::Convert<PdfNewPage>(page)->GetAnnotations()->Add(annotation);
		y += size->GetHeight();
		return y;
	}

	float AddPopupAnnotation(intrusive_ptr<PdfPageBase> page, float y)
	{
		intrusive_ptr<PdfTrueTypeFont> font = new PdfTrueTypeFont(L"Arial", 12.f, PdfFontStyle::Regular, true);
		intrusive_ptr<PdfStringFormat> format = new PdfStringFormat();
		format->SetMeasureTrailingSpaces(true);
		wstring prompt = L"Markup incorrect spelling: ";
		intrusive_ptr<SizeF> size = font->MeasureString(prompt.c_str(), format);
		page->GetCanvas()->DrawString(prompt.c_str(), font, PdfBrushes::GetDodgerBlue(), 0, y);
		float x = size->GetWidth();
		wstring label = L"demo of anotation";
		page->GetCanvas()->DrawString(label.c_str(), font, PdfBrushes::GetOrangeRed(), x, y);
		x += font->MeasureString(label.c_str(), format)->GetWidth();
		wstring markupText = L"All words were spelled correctly";
		size = font->MeasureString(markupText.c_str());
		intrusive_ptr<PdfPopupAnnotation> annotation
			= new PdfPopupAnnotation(new RectangleF(new PointF(x, y), SizeF::Empty()), markupText.c_str());
		annotation->SetIcon(PdfPopupIcon::Paragraph);
		annotation->SetOpen(true);
		annotation->SetColor(new PdfRGBColor(Spire::Common::Color::GetYellow()));
		Object::Convert<PdfNewPage>(page)->GetAnnotations()->Add(annotation);
		y += size->GetHeight();
		return y;
	}

	float AddRubberStampAnnotation(intrusive_ptr<PdfPageBase> page, float y)
	{

		intrusive_ptr<PdfTrueTypeFont> font = new PdfTrueTypeFont(L"Arial", 12.f, PdfFontStyle::Bold, true);
		intrusive_ptr<PdfStringFormat> format = new PdfStringFormat();
		format->SetMeasureTrailingSpaces(true);
		wstring prompt = L"Markup incorrect spelling: ";
		intrusive_ptr<SizeF> size = font->MeasureString(prompt.c_str(), format);
		page->GetCanvas()->DrawString(prompt.c_str(), font, PdfBrushes::GetDodgerBlue(), 0, y);
		float x = size->GetWidth();
		wstring label = L"demo of annotation";
		page->GetCanvas()->DrawString(label.c_str(), font, PdfBrushes::GetOrangeRed(), x, y);
		x += font->MeasureString(label.c_str(), format)->GetWidth();
		wstring markupText = L"Just a draft, not checked.";
		size = font->MeasureString(markupText.c_str());
		intrusive_ptr<PdfRubberStampAnnotation> annotation
			= new PdfRubberStampAnnotation(new RectangleF(x, y, font->GetHeight(), font->GetHeight()), markupText.c_str());
		annotation->SetIcon(PdfRubberStampAnnotationIcon::Draft);
		annotation->SetColor(new PdfRGBColor(Spire::Common::Color::GetPlum()));
		Object::Convert<PdfNewPage>(page)->GetAnnotations()->Add(annotation);
		y += size->GetHeight();
		return y;
	}
	
	int main() {
		wstring output_path = OUTPUTPATH;
		wstring outputFile = output_path + L"Annotation.pdf";
		intrusive_ptr<PdfDocument> doc = new PdfDocument();
		intrusive_ptr<PdfUnitConvertor> unitCvtr = new PdfUnitConvertor();
		intrusive_ptr<PdfMargins> margin = new PdfMargins();
		margin->SetTop(unitCvtr->ConvertUnits(2.54f, PdfGraphicsUnit::Centimeter, PdfGraphicsUnit::Point));
		margin->SetBottom(margin->GetTop());
		margin->SetLeft(unitCvtr->ConvertUnits(3.f, PdfGraphicsUnit::Centimeter, PdfGraphicsUnit::Point));
		margin->SetRight(margin->GetLeft());
		intrusive_ptr<PdfPageBase> page = doc->GetPages()->Add(new SizeF(PdfPageSize::A4()), margin);
		intrusive_ptr<PdfBrush> brush1 = PdfBrushes::GetBlack();
		intrusive_ptr<PdfTrueTypeFont> font1 = new PdfTrueTypeFont(L"Arial", 13.f, PdfFontStyle::Bold | PdfFontStyle::Italic, true);
		intrusive_ptr<PdfStringFormat> format1 = new PdfStringFormat(PdfTextAlignment::Left);
		float y = 50;
		wstring s = L"The sample demonstrates how to add annotations in PDF document.";
		page->GetCanvas()->DrawString(s.c_str(), font1, brush1, 0, y - 5, format1);
		y += font1->MeasureString(s.c_str(), format1)->GetHeight();
		y = y + 15;


		y = AddDocumentLinkAnnotation(page, y);

		y = y + 6;
		y = AddFileLinkAnnotation(page, y);

		y = y + 6;
		y = AddFreeTextAnnotation(page, y);

		y = y + 6;
		y = AddLineAnnotation(page, y);

		y = y + 6;
		y = AddTextMarkupAnnotation(page, y);

		y = y + 6;
		y = AddPopupAnnotation(page, y);

		y = y + 6;
		y = AddRubberStampAnnotation(page, y);

		//Save the document
		doc->SaveToFile(outputFile.c_str(), FileFormat::PDF);
		doc->Close();
	}


