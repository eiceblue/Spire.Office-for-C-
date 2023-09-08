#include "pch.h"

using namespace std;
using namespace Spire::Pdf;
using namespace Spire::Common;


void DrawTableInHeaderFooter(intrusive_ptr<PdfDocument> doc)
{
	std::vector<LPCWSTR_S> data = { L"Column1;Column2", L"Spire.PDF for .NET;Spire.PDF for JAVA" };
	float y = 20;
	intrusive_ptr<PdfBrush> brush = PdfBrushes::GetBlack();

	for (int i = 0; i < doc->GetPages()->GetCount(); i++)
	{
		intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(i);
		//Create Pdf table
		intrusive_ptr<PdfTable> table = new PdfTable();
		table->GetStyle()->SetCellPadding(2);
		table->GetStyle()->SetBorderPen(new PdfPen(brush, 0.1f));
		table->GetStyle()->GetHeaderStyle()->SetStringFormat(new PdfStringFormat(PdfTextAlignment::Center));
		table->GetStyle()->SetHeaderSource(PdfHeaderSource::Rows);
		table->GetStyle()->SetHeaderRowCount(1);
		table->GetStyle()->SetShowHeader(true);
		table->GetStyle()->GetHeaderStyle()->SetBackgroundBrush(PdfBrushes::GetCadetBlue());

		table->SetDataSource(data);

		intrusive_ptr<PdfColumn> column = new PdfColumn();

		for (int j = 0; j < table->GetColumns()->GetCount(); j++)
		{
			column = table->GetColumns()->GetItem(j);
			column->SetStringFormat(new PdfStringFormat(PdfTextAlignment::Center, PdfVerticalAlignment::Middle));
		}

		//Draw the table on page
	
		intrusive_ptr<PdfLayoutWidget> layoutTable = Object::Convert<PdfLayoutWidget>(table);
		layoutTable->Draw(page, new PointF(0, y));
	}
}

int main()
{
	wstring inputFile = DATAPATH"PDFTemplate_HF.pdf";
	wstring outputFile = OUTPUTPATH"TableInHeaderFooter.pdf";

	//Load the Pdf from disk
	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());

	//Draw table in header
	DrawTableInHeaderFooter(doc);

	//Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}

