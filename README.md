# AlgMeter
Simplest way of measuring an algorithm execution time and generating graphs

# Usage
AlgMeter is composed of two main modules: Meter and Renderer. Meter measures an algorithm's execution time and renderer generates a spreadsheet with a table and graph of the times measured by Meter. Its usage is as simple as follows:

```java
Meter meter = new Meter(n -> new Foo().execute(n), startN, endN, 
    previousN -> previousN * 2, repetitions);

Renderer renderer = new Renderer(meter.run(), //Map<Long, Long> with execution times
                                 "algtimes"); //generated spreadsheet name (without extension)
renderer.render();
```
