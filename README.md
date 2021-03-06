# PowerPoint Editor
A library to used to provide basic editing capabilities of PowerPoint presentation. It is primarily going to be used to give non-coders this ability via .net Interactive Notebooks. However, it could easily be used .net apps as well

[![Nuget](https://img.shields.io/nuget/v/Fritz.InstantAPIs)](https://www.nuget.org/packages/CFXYZ.PowerPointEditorLib/)
[![Instant APIs Documentation](https://img.shields.io/badge/docs-ready!-blue)](https://github.com/cfelstead/PowerPointEditor)
![GitHub last commit](https://img.shields.io/github/last-commit/cfelstead/PowerPointEditor)

## Examples

### Setup

var presentation = new PowerPointPresentation(pathToThePptx);

### Text replacement

```csharp
presentation.ForAllSlides()
            .ReplaceText("{{TEAM_NAME}}")
            .With("FooBar Unitied");

presentation.ForSlide(3)
            .ReplaceText("{{TEAM_NAME}}")
            .With("FooBar Unitied");
```

As you can see from the above, you can either specify to work with one single slide by its position in the deck with `ForSlide(3)` or all slides with `ForAllSlides()`.

Inside the powerpoint you are using as a template, we are looking for the text ***{{TEAM_NAME}}*** and will replace it with ***FooBar United***.

Sometime the text can be found in different sections of the powerpoint xml. This is because the styling/formatting changes mid sentence/paragraph. In these circumstances, you can automatically remove all the formatting changes (leaving it as the initial section is) and then replace the text. To do this use `IgnoringStylingWith` instead of `With`.

```csharp
presentation.ForAllSlides()
            .ReplaceText("{{TEAM_NAME}}")
            .IgnoringStylingWith("FooBar Unitied");

presentation.ForSlide(3)
            .ReplaceText("{{TEAM_NAME}}")
            .IgnoringStylingWith("FooBar Unitied");
```

### Working with images

```csharp
presentation.ForSlide(2)
            .FindPictureWithAltText("TeamLogo")
            .ReplaceImageWith("C:\MyTeamLogo.jpg");

presentation.ForAllSlides()
            .FindPictureWithAltText("TeamLogo")
            .ReplaceImageWith("C:\MyTeamLogo.jpg");

presentation.ForSlide(2)
            .FindPictureWithAltText("TeamLogo")
            .ReplaceAltTextWith($"FooBar United Team Logo");
```

Here we are looking in slide 2 for an image with the alternate text to ***TeamLogo***. We then replace the image with a jpg from our local disk and change the alternate text to be ***FooBar United Team Logo***.

### Changing image hyperlinks

```csharp
presentation.ForSlide(2)
            .FindPictureWithAltText("TeamLogo")
            .ReplaceHyperlinkWith("http://www.newlink.com");
```

Here we are looking in slide 2 for an image with the alternate text to ***TeamLogo***. We then replace an existing hyperlink (there should already be one) with a new url.

### Working with bar charts

```csharp
presentation.ForSlide(3)
            .FindBarChart(1)
            .FindSeriesByName("Category A")
            .ReplaceValuesWith(new List<string> { "0.15", "0.3", "0.2", "0.2" });

presentation.ForSlide(3)
            .FindBarChart(1)
            .FindSeriesByName("Category B")
            .ReplaceValuesWith(new List<string> { "0.2", "0.4", "0.3", "0.32" });

presentation.ForSlide(3)
            .FindBarChart(1)
            .FindSeriesByName("Category C")
            .ReplaceValuesWith(new List<string> { "0.4", "0.8", "0.5", "0.48" });
```

We are editing the first bar chart on slide 3. It has 3 series and 4 values per series. We are adjusting the values of the chart to the percentages we have passed in.

### Working with pie charts

```csharp
presentation.ForSlide(4)
            .FindPieChart(2)
            .FindSeries()
            .ReplaceValuesWith(new List<string> { "0.2", "0.25", "0.1", "0.45" });
```

We are editing the second pie chart on slide 4. It has 4 values and we are adjusting them to the values we have passed in.

### Working with table contents

```csharp
presentation.ForSlide(5)
            .FindTable(1)
            .FindRow(2)
            .FindCell(1)
            .ReplaceCellText("Default cell text", "New cell text");
```

We are editing the first table on slide 5. We are then editing the second row and first column to replace the text "Default cell text" with "New cell text".