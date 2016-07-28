package viewController;

import javafx.application.Application;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.layout.FlowPane;
import javafx.scene.layout.StackPane;
import javafx.stage.Stage;
import model.Tester;

import java.io.IOException;

public class Main extends Application {


    @Override
    public void start(Stage primaryStage) throws Exception{
        primaryStage.setTitle("Hello World!");
        Button btn = new Button();
        btn.setText("'read wb'");
        btn.setOnAction( (ActionEvent event) -> {
            try {
                Tester.readWb();
            } catch (IOException e) {
                e.printStackTrace();
            }
        });

        Button btn2 = new Button();
        btn2.setText("'write wb'");
        btn2.setOnAction( (ActionEvent event)->{
            try {
                Tester.createWb();
            } catch (IOException e) {
                e.printStackTrace();
            }
        });

        StackPane root = new StackPane();
        FlowPane flow = new FlowPane();
        flow.setPadding(new Insets(5, 0, 5, 0));
        flow.setVgap(4);
        flow.setHgap(4);
        flow.setPrefWrapLength(170); // preferred width allows for two columns
        flow.setStyle("-fx-background-color: DAE6F3;");
        root.getChildren().add(flow);
        flow.getChildren().add(btn2);
        flow.getChildren().add(btn);
        primaryStage.setScene(new Scene(root, 300, 250));
        primaryStage.show();
    }


    public static void main(String[] args) {
        launch(args);
    }
}
