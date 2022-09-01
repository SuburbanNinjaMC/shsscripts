public class MyClass {
    public static void permutations(String str) {
        permutations(str.toCharArray(), 0, 0);
    }
    
    public static void permutations(char[] list, int pivot, int start) {
        if (pivot == list.length) {
            System.out.println(charArray(list, 0, list.length));
            return;
        }
        
        if (start == list.length) {
            return;
        }
        
        swap(list, pivot, start);
        permutations(list, pivot + 1, start);
        swap(list, pivot, start);
        
        permutations(list, pivot, start + 1);
    }
    
    public static String charArray(char[] vals, int start, int end) {
        if (start == end) return "";
        return vals[start] + charArray(vals, start + 1, end);
    }
    
    public static void swap(char[] arr, int aIndex, int bIndex) {
        char temp = arr[aIndex];
        arr[aIndex] = arr[bIndex];
        arr[bIndex] = temp;
    }
}
