package mainPKG.toolRelationship;

import java.io.PrintStream;
import java.util.List;

public enum DotPrinter {
	INSTANCE;
	
	private PrintStream printStream = System.out;
	private int indent = 0;
	private int indentAtom = 1;
	private void indentAdd() {
		indent += indentAtom;
	}
	private void indentSub() {
		indent -= indentAtom;
	}
	public <T> void printDots(List<Dot<T>> dots, PrintStream printStream) {
		if (printStream != null) {
			this.printStream = printStream;
		}
		this.printStream.println("--------------------------------------------------------------------");
		for (Dot<T> dot : dots) {
			printDot(dot, null);
		}
	}
	public <T> void printDot(Dot<T> rootDot, PrintStream printStream) {
		indent = 0;
		if (printStream != null) {
			this.printStream = printStream;
		}
		for (int i = 0; i < 12; i++) {
			this.printStream.print(i);
			this.printStream.print('\t');
		}
		this.printStream.println();
		toPrint(rootDot);
		this.printStream.println("--------------------------------------------------------------------");
	}
	private <T> void toPrint(Dot<T> rootDot) {
		printInfo(rootDot.getCurrPro().toString());
		if (rootDot.hasChildrenDots()) {
			indentAdd();
			printNewLine();
			for (Dot<T> dotChild : rootDot.getChildrenDots()) {
				printIndent();
				toPrint(dotChild);
			}
			indentSub();
		} else {
			printNewLine();
		}
	}
	private void printInfo(String info) {
		printStream.print(info);
	}
	private void printIndent() {
		for (int i = 0; i < indent; i++) {
			printStream.print("|\t");
		}
	}
	private void printNewLine() {
		printStream.println();
	}
}
