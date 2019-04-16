/**
 * 
 */
package mainPKG.toolRelationship;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

/**
 * 调用关系解析
 */
public class CallRelationshipAnalyze<T> {
	
	/**存储了初始调用关系的hashMap的引用*/
	private Map<T, List<T>> hashMap = null;
	
	/**存儲用於返回ArrayList序列*/
	private List<Dot<T>> childrenDots = null;
	public List<Dot<T>> getAllChldrenDots() {
		return childrenDots;
	}
	private List<Dot<T>> rootDots = null;
	public List<Dot<T>> getAllRootDots() {
		return rootDots;
	}
	
	/**
	 * <p>從給定的hashMap中解析调用关系。</p>
	 * 
	 * @param	newHashMap			存储了初始调用关系的hashMap
	 * @param	addMore				true為重複添加模式，false為不重複添加模式
	 */
	public void toResolve(Map<T, List<T>> newHashMap, boolean addMore) {
		hashMap =newHashMap;
		childrenDots = new ArrayList();
		rootDots = new ArrayList();
		
		Dot<T> rootDotTmp = null;
		//有重複添加
		if (addMore == true) {
			//遍歷hashMap
			for (Entry<T, List<T>> entry : hashMap.entrySet()) {
				//主函數情形
				if (!mapContainsValue(entry.getKey())) {
					T startMainPro = entry.getKey();
					rootDotTmp = new Dot<T>(null, startMainPro, 0);
					rootDots.add(rootDotTmp);
					trueAddDot(rootDotTmp);
				}
			}
		} else {//無重複添加
			//遍歷hashMap
			for (Entry<T, List<T>> entry : hashMap.entrySet()) {
				//主函數情形
				if (!mapContainsValue(entry.getKey())) {
					T startMainPro = entry.getKey();
					rootDotTmp = new Dot<T>(null, startMainPro, 0);
					rootDots.add(rootDotTmp);
					falseAddDot(rootDotTmp);
				}
			}
		}
		
	}
	
	/**
	 * <p>遞歸調用函數。它在arrayList 中(有重複)添加調用關係對象節點。</p>
	 * 
	 * @param	lastDot						遞歸過程中的主調函數節點
	 */
	private void trueAddDot(Dot<T> lastDot) {
		Dot<T> dotTmp;
		//遍歷所有被調函數(list)
		for (T subProTmp : hashMap.get(lastDot.getCurrPro())) {
			//創建備用節點
			dotTmp = new Dot<T>(lastDot.getMainProDot(), subProTmp, lastDot.getGrade()+1);
			childrenDots.add(dotTmp);//創建并添加
			//設置前後關聯
			lastDot.addChildDot(dotTmp);
			dotTmp.setFatherDot(lastDot);
			
			//有後續調用，且不出現調用循環，則繼續
			if (hashMap.containsKey(subProTmp) && !isAdded(dotTmp)){
				trueAddDot(dotTmp);//遞歸遞歸添加
			}
		}
	}
	
	/**
	 * <p>判斷同一調用鏈條中參考節點是否被添加過。</p>
	 * 
	 * @param	dotTmp		用於判斷的參考節點
	 * @return	boolean		true添加過，false未添加過
	 */
	private boolean isAdded(Dot<T> dotTmp) {
		Dot<T> dot = dotTmp.getFatherDot();//存儲索引變量
		//在調用鏈條中循環查找
		while (dot != null) {
			if (dotTmp.getMainProDot().equals(dot.getMainProDot()) && dotTmp.getCurrPro().equals(dot.getCurrPro())) {
				return true;
			} else {
				dot = dot.getFatherDot();//繼續下一判斷
			}
		}
		//此時必然未添加
		return false;
	}

	/**
	 * <p>遞歸調用函數。它在arrayList 中(無重複)添加調用關係對象節點。</p>
	 * 
	 * @param	lastDot						遞歸過程中的主調函數節點
	 */
	private void falseAddDot(Dot<T> lastDot) {
		Dot<T> dotTmp;
		//遍歷所有被調函數(list)
		for (T subProTmp : hashMap.get(lastDot.getCurrPro())) {
			//創建備用節點
			dotTmp = new Dot<T>(lastDot.getMainProDot(), subProTmp, lastDot.getGrade()+1);
			//判斷后添加
			if (!shouldAdd(dotTmp)) {
				continue;
			}
			childrenDots.add(dotTmp);//創建并添加
			//設置前後關聯
			lastDot.addChildDot(dotTmp);
			dotTmp.setFatherDot(lastDot);
			
			//有後續調用，則繼續
			if (hashMap.containsKey(subProTmp)){
				falseAddDot(dotTmp);//遞歸遞歸添加
			}
		}
	}
	
	/**
	 * <p>判斷hashMap中的TALUE列(List)中是否包含參數字符串。</p>
	 * 
	 * @param	mainPro				用於判斷的參數字符串
	 * @return	boolean				true包含，false不包含
	 */
	private boolean mapContainsValue(T mainPro) {
		//空值檢查
		if (hashMap==null || hashMap.isEmpty()) {
			return false;
			
		}
		//遍歷hashMap中的string列
		for(T key : hashMap.keySet()) {
			if (hashMap.get(key).contains(mainPro)) {
				return true;
			}
		}
		//此時必然不包含
		return false;
	}
	
	/**
	 * <p>不可重複添加時，判斷arraylist 中是否應添加該節點。</p>
	 * @param	compareDot				遞歸過程中的主調函數節點
	 * @return	boolean				true應添加，false不應添加
	 */
	private boolean shouldAdd(Dot<T> compareDot) {
		//arrayList空值判斷
		if (childrenDots.isEmpty()) {
			return true;
		}
		//取出數據
		Dot<T> mainProDot = compareDot.getMainProDot();
		T subPro = compareDot.getCurrPro();
		int grade = compareDot.getGrade();
		//遍歷arraylist
		for (Dot<T> dotOld : childrenDots) {
			//比較
			if (mainProDot.equals(dotOld.getMainProDot()) && subPro.equals(dotOld.getCurrPro())) {//有同名記錄
				if (grade>=dotOld.getGrade()) {//記錄中的調用層級數相等或更小
					return false;//保持記錄，不應添加
				} else {//記錄中的調用層級數更高
					childrenDots.remove(dotOld);//移除后
					return true;//再添加
				}
			}
		}
		//此時必然應添加
		return true;
	}
	
}
