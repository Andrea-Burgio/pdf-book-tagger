package org.SchedulesExtractor;


public class Main {
    public static void main(String[] args) {
            String XMlClassificationPath = args[0];
            System.out.println("Creating schedules from: "  + XMlClassificationPath + "...");
            SchedulesExtractor.extractSchedules(XMlClassificationPath);
            System.out.println("\u001B[38;2;79;196;20mSchedules extracted successfully.\u001B[0m");
    }
}