package mainPKG.toolRelationship;

import java.util.ArrayList;
import java.util.List;

/**
 * 自定義用於存儲調用關係與調用層級數相關類
 */

public class Dot<T> {
	
	private Object hook = null;
	public Object getHook() {
		return hook;
	}
	public void setHook(Object hook) {
		this.hook = hook;
	}
	
	/** 主程序名 */
	private Dot<T> mainPro = null;
	/** 子程序名 */
	private T currPro = null;
	/** 調用級數 */
	private int grade = 0;
	/**上一節點的索引*/
	private Dot<T> fatherDot =null;
	/**下一節點的索引*/
	private List<Dot<T>> childrenDots =null;
	
	public boolean hasChildrenDots() {
		return childrenDots != null;
	}
	public boolean isRootDot() {
		return mainPro == null;
	}

	/**
	 * <p>構造方法</p>
	 * @param	mainPro	主程序名
	 * @param	subPro	子程序名
	 * @param	grade	調用級數
	 */
	public Dot(Dot<T> mainPro, T subPro, int grade) {
		this.mainPro = mainPro;
		this.currPro = subPro;
		this.grade = grade;
	}
	/**
	 * @return 主程序名
	 */
	public Dot<T> getMainProDot() {
		return mainPro==null ? this : mainPro;
	}
	/**
	 * @return 子程序名
	 */
	public T getCurrPro() {
		return currPro;
	}
	/**
	 * @return 調用級數
	 */
	public int getGrade() {
		return grade;
	}
	/**
	 * @return 上一節點
	 */
	public Dot<T> getFatherDot() {
		return fatherDot;
	}
	/**
	 * @param fatherDot 上一節點設置
	 */
	public void setFatherDot(Dot<T> fatherDot) {
		this.fatherDot = fatherDot;
	}
	/**
	 * @return 所有子节点
	 */
	public List<Dot<T>> getChildrenDots() {
		return  childrenDots;
	}
	/**
	 * @param nextDot 子节点
	 */
	public void addChildDot(Dot<T> nextDot) {
		if (childrenDots == null) {
			childrenDots = new ArrayList();
		}
		childrenDots.add(nextDot);
	}
	
	/**
	 * <p>用於跟參考節點進行比較。</p>
	 * @param	compareDot	用於比較的參考節點
	 * @return	true相等，false不相等
	 */
	public boolean equalsThisDot(Dot<T> compareDot) {
		return compareDot.getMainProDot().equals(mainPro) && compareDot.getCurrPro().equals(currPro) && compareDot.getGrade()==grade;
	}
	
	@Override
	public String toString() {
		return "["+mainPro.getCurrPro().toString()+":"+currPro.toString()+":"+grade+"]";
	}
	
}
