package mainPKG.framework2File.taskModel;

import java.io.File;

public interface WorkFileTask {

	public void task_Init(File file, int index, int allCnt);
	
	public void task_ToDo(File file, int index, int allCnt);
	
	public void task_End(File file, int index, int allCnt);
	
}
