package mainPKG.framework2File.fileModel;

public class BasicSrcLine {

	protected String srcContent;
	
	protected BasicSrcLine prevLine;
	protected BasicSrcLine nextLine;
	
	public BasicSrcLine(String _srcContent) {
		srcContent = _srcContent;
	}
	public BasicSrcLine(BasicSrcLine prevLine) {
		if (prevLine.nextLine != null) {
			this.setNextLine(prevLine.nextLine);
			prevLine.nextLine.setPrevLine(this);
		}
		this.setPrevLine( prevLine );
		prevLine.setNextLine( this );
	}
	
	public String getSrcContent() {
		return srcContent;
	}
	public void setSrcContent(String srcContent) {
		this.srcContent = srcContent;
	}

	public BasicSrcLine getPrevLine() {
		return prevLine;
	}
	public void setPrevLine(BasicSrcLine prevLine) {
		this.prevLine = prevLine;
	}
	public BasicSrcLine getNextLine() {
		return nextLine;
	}
	public void setNextLine(BasicSrcLine nextLine) {
		this.nextLine = nextLine;
	}
	
	
	@Override
	public String toString() {
		return srcContent;
	}
	
}
