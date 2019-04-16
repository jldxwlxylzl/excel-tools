package mainPKG.framework2File.fileModel;

public class SrcIndexLine extends BasicSrcLine {
	
	public SrcIndexLine(String srcContent) {
		super(srcContent);
	}
	public SrcIndexLine(SrcIndexLine prevLine) {
		super(prevLine);
	}
	
	protected int lineIndex = -1;
	public void setLineIndex(int _lineIndex, boolean affectAfterLines) {
		lineIndex = _lineIndex;
		if (affectAfterLines && nextLine!=null) {
			((SrcIndexLine) nextLine).setLineIndex(lineIndex+1, affectAfterLines);
		}
	}
	public int getLineIndex() {
		return lineIndex;
	}
	
}
