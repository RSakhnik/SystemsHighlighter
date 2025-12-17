using System;
using System.Collections.Generic;
using System.Windows.Controls;
using System.Windows;

namespace SystemsHighlighter.Tools
{
    public class DateSlider : Slider
    {
        public static readonly DependencyProperty DateLabelsProperty = DependencyProperty.Register(
            nameof(DateLabels),
            typeof(List<DateTime>),
            typeof(DateSlider),
            new PropertyMetadata(null));

        public List<DateTime> DateLabels
        {
            get => (List<DateTime>)GetValue(DateLabelsProperty);
            set => SetValue(DateLabelsProperty, value);
        }

        // Убираем переопределение стиля, чтобы использовать стандартный шаблон Slider
    }
}