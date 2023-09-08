#include "pch.h"

using namespace Spire::Presentation;

int main()
{

	wstring outputFile = OUTPUTPATH"CloneRowAndColumn.pptx";

	intrusive_ptr<Presentation> presentation = new Presentation();
	// Access first slide
	intrusive_ptr<ISlide> sld = presentation->GetSlides()->GetItem(0);

	// Define columns with widths and rows with heights
	vector<double> widths = { 110, 110, 110 };
	vector<double> heights = { 50, 30, 30, 30, 30 };

	// Add table shape to slide
	intrusive_ptr<ITable> table = presentation->GetSlides()->GetItem(0)->GetShapes()->AppendTable(presentation->GetSlideSize()->GetSize()->GetWidth() / 2 - 275, 90, widths, heights);

	// Add text to the row 1 cell 1
	table->GetItem(0, 0)->GetTextFrame()->SetText(L"Row 1 Cell 1");

	// Add text to the row 1 cell 2
	table->GetItem(1, 0)->GetTextFrame()->SetText(L"Row 1 Cell 2");

	// Clone row 1 at end of table
	table->GetTableRows()->Append(table->GetTableRows()->GetItem(0));

	// Add text to the row 2 cell 1
	table->GetItem(0, 1)->GetTextFrame()->SetText(L"Row 2 Cell 1");

	// Add text to the row 2 cell 2
	table->GetItem(1, 1)->GetTextFrame()->SetText(L"Row 2 Cell 2");

	// Clone row 2 as the 4th row of table
	table->GetTableRows()->Insert(3, table->GetTableRows()->GetItem(1));

	//Clone column 1 at end of table
	table->GetColumnsList()->Add(table->GetColumnsList()->GetItem(0));

	//Clone the 2nd column at 4th column index
	table->GetColumnsList()->Insert(3, table->GetColumnsList()->GetItem(1));

	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	
}
