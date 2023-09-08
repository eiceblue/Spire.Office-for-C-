#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Word.png";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetOutsidePosition.docx";

	//Create a new word document and add new section
	intrusive_ptr<Document> doc = new Document();
	intrusive_ptr<Section> sec = doc->AddSection();

	//Get header
	intrusive_ptr<HeaderFooter> header = doc->GetSections()->GetItemInSectionCollection(0)->GetHeadersFooters()->GetHeader();

	//Add new paragraph on header and set HorizontalAlignment of the paragraph as left
	intrusive_ptr<Paragraph> paragraph = header->AddParagraph();
	paragraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Left);

	//Load an image for the paragraph
	intrusive_ptr<DocPicture> headerimage = paragraph->AppendPicture(inputFile.c_str());
	//Add a table of 4 rows and 2 columns
	intrusive_ptr<Table> table = header->AddTable();
	table->ResetCells(4, 2);

	//Set the position of the table to the right of the image
	table->GetTableFormat()->SetWrapTextAround(true);
	table->GetTableFormat()->GetPositioning()->SetHorizPositionAbs(HorizontalPosition::Outside);
	table->GetTableFormat()->GetPositioning()->SetVertRelationTo(VerticalRelation::Margin);
	table->GetTableFormat()->GetPositioning()->SetVertPosition(43);

	//Add contents for the table
	std::vector<std::vector<std::wstring>> data =
	{
		{L"Spire.Doc.left", L"Spire XLS.right"},
		{L"Spire.Presentatio.left", L"Spire.PDF.right"},
		{L"Spire.DataExport.left", L"Spire.PDFViewe.right"},
		{L"Spire.DocViewer.left", L"Spire.BarCode.right"}
	};

	for (int r = 0; r < 4; r++)
	{
		intrusive_ptr<TableRow> dataRow = table->GetRows()->GetItemInRowCollection(r);
		for (int c = 0; c < 2; c++)
		{
			if (c == 0)
			{
				intrusive_ptr<Paragraph> par = dataRow->GetCells()->GetItemInCellCollection(c)->AddParagraph();
				par->AppendText(data[r][c].c_str());
				par->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Left);
				dataRow->GetCells()->GetItemInCellCollection(c)->SetWidth(180);
			}
			else
			{
				intrusive_ptr<Paragraph> par = dataRow->GetCells()->GetItemInCellCollection(c)->AddParagraph();
				par->AppendText(data[r][c].c_str());
				par->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Right);
				dataRow->GetCells()->GetItemInCellCollection(c)->SetWidth(180);
			}
		}
	}

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}
