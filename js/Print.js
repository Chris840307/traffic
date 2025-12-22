function printWindow(portraitVal,leftMarginVal,topMarginVal,rightMarginVal,bottomMargin) {
factory.printing.header = "";
factory.printing.footer = "";
factory.printing.portrait = portraitVal;
factory.printing.leftMargin = leftMarginVal;
factory.printing.topMargin = topMarginVal;
factory.printing.rightMargin = rightMarginVal;
factory.printing.bottomMargin = bottomMargin;
factory.printing.Print(true);
return true;
}

function AutoprintWindow(portraitVal,leftMarginVal,topMarginVal,rightMarginVal,bottomMargin) {
factory.printing.header = "";
factory.printing.footer = "";
factory.printing.portrait = portraitVal;
factory.printing.leftMargin = leftMarginVal;
factory.printing.topMargin = topMarginVal;
factory.printing.rightMargin = rightMarginVal;
factory.printing.bottomMargin = bottomMargin;
factory.printing.Print(false);
return true;
}