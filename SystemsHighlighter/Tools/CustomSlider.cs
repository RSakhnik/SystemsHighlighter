using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;

namespace SystemsHighlighter.Tools
{
    public class CustomSlider : Slider
    {
        static CustomSlider()
        {
            // Чтобы WPF применял для нашего контролла тот же шаблон, что и для обычного Slider:
            DefaultStyleKeyProperty.OverrideMetadata(
                typeof(CustomSlider),
                new FrameworkPropertyMetadata(typeof(Slider)));

            // Отключаем отображение штрихов (тиков) под слайдером
            TickPlacementProperty.OverrideMetadata(
                typeof(CustomSlider),
                new FrameworkPropertyMetadata(TickPlacement.None));

            // Отключаем всплывающие подсказки с текущим значением при тащении ползунка
            AutoToolTipPlacementProperty.OverrideMetadata(
                typeof(CustomSlider),
                new FrameworkPropertyMetadata(AutoToolTipPlacement.None));
        }

        public CustomSlider()
        {
            // Snap to ticks можно оставить, если нужно фиксированное количество шагов:
            this.IsSnapToTickEnabled = true;
            this.TickFrequency = 1;

            // Но сами тики и их подписи не будут отображаться
        }
    }
}
