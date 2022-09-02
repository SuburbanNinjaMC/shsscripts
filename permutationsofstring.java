public class Permutations {
    // Discovered this problem online searching for "difficult recursion practice." Was a fun challenge.
    
    public static void permutations(String str) {
        permutations(str.toCharArray(), 0, 0);
    }
    
    public static void permutations(char[] list, int index, int start) {
        // Once the index has reached the end of the array, we're done.
        if (index == list.length - 1) {
            System.out.println(charArray(list, 0, list.length));
            return;
        }

        // Repeated swaps that allow us to randomize the array. We keep the first letter constant and then recursively permute the rest with the same logic.
        // Attempted to make this completely recursive. Failed. Hybrid solution implemented instead.
        start = index;
        while (start < list.length) {
            swap(list, index, start);
            // Recursive call.
            permutations(list, index + 1, start);
            
            // When we swap it back, we wind up on the same letter.
            swap(list, index, start);
            start++;
        }
    }

    //  While the String.valueOf(char[] T) exists, I decided to go full recursive here.
    public static String charArray(char[] vals, int start, int end) {
        if (start == end) return "";
        return vals[start] + charArray(vals, start + 1, end);
    }
    
    // Basic swap method, learned from my old computer science course.
    public static void swap(char[] arr, int aIndex, int bIndex) {
        char temp = arr[aIndex];
        arr[aIndex] = arr[bIndex];
        arr[bIndex] = temp;
    }
    
    // Tester method.
    public static void main(String[] args) {
        permutations("abcd");
    }
}
