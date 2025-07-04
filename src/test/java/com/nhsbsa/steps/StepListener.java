package com.nhsbsa.steps;
import io.cucumber.plugin.ConcurrentEventListener;
import io.cucumber.plugin.event.EventHandler;
import io.cucumber.plugin.event.EventPublisher;
import io.cucumber.plugin.event.PickleStepTestStep;
import io.cucumber.plugin.event.TestStepStarted;
public class StepListener implements ConcurrentEventListener {
                // public static String stepName;
                public static ThreadLocal<String> stepName = new ThreadLocal<String>();
                public EventHandler<TestStepStarted> stepHandler = new EventHandler<TestStepStarted>() {
                                @Override
                                public void receive(TestStepStarted event) {
                                                handleTestStepStarted(event);
                                }
                };
                @Override
                public void setEventPublisher(EventPublisher publisher) {
                                publisher.registerHandlerFor(TestStepStarted.class, stepHandler);
                }
                private void handleTestStepStarted(TestStepStarted event) {
                                if (event.getTestStep() instanceof PickleStepTestStep) {
                                                PickleStepTestStep testStep = (PickleStepTestStep) event.getTestStep();
                                                stepName.set(testStep.getStep().getText());
                                                // stepStatus.set(testStep.);
                                }
                }
}