import java.io.File;
import java.util.Scanner;

public class results {
    public static void main(String[] args) {
        start();
    }

    public static void start() {
        Scanner sc = new Scanner(System.in);
        System.out.print("请输入总成绩表路径:");
        String file = sc.nextLine();
        System.out.print("分科成绩将会保存在总成绩表同路径下。\n运行中......\n");
        new method(new File(file.trim()), new method.End() {
            @Override
            public void end() {
                System.out.print("本次任务已完成。\n");
                start();
            }
        }).start();
    }
}
