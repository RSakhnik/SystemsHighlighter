using Ascon.Pilot.Bim.SDK;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SystemsHighlighter.Tools
{
    public class SystemsData
    {
        // Элемент с Guid и Pipeline
        public class Element
        {
            public ForColorModelElementId ColorId { get; }
            public string Weight { get; }
            public string Lenght { get; }
            public string Volume { get; }
            public string DiaInch { get; }

            public Element(Guid elementId, string section, string weight, string lenght, string volume, string diaInch, Dictionary<string, Guid> SectionMappings)
            {
                Guid modelPartId;
                if (!SectionMappings.TryGetValue(section, out modelPartId))
                    modelPartId = Guid.Parse("");

                ColorId = new ForColorModelElementId(elementId, modelPartId);
                Weight = weight;
                Lenght = lenght;
                Volume = volume;
                DiaInch = diaInch;
            }
        }

        // Трубопроводная линия: имя + список элементов
        public class PipeLine
        {
            public string Name { get; set; }
            public List<Element> Elements { get; set; } = new List<Element>();

            public PipeLine(string name)
            {
                Name = name ?? throw new ArgumentNullException(nameof(name));
            }

            public void AddElement(Guid elementId, string section, string weight, string length, string volume, string diainch, Dictionary<string, Guid> SectionMappings)
            {
                var elem = new Element(elementId, section, weight, length, volume, diainch, SectionMappings);
                Elements.Add(elem);
            }
        }

        // Подсистема: имя + список элементов
        public class Subsystem
        {
            public string Name { get; }
            public List<PipeLine> PipeLines { get; set; } = new List<PipeLine>();

            public Subsystem(string name)
            {
                Name = name ?? throw new ArgumentNullException(nameof(name));
            }

            public PipeLine GetOrCreatePipeLine(string pipeLineName)
            {
                var pipeLine = PipeLines.FirstOrDefault(p => p.Name == pipeLineName);
                if (pipeLine == null)
                {
                    pipeLine = new PipeLine(pipeLineName);
                    PipeLines.Add(pipeLine);
                }
                return pipeLine;
            }
        }

        // Система: имя + список подсистем
        public class SystemClass
        {
            public string Name { get; }
            public List<Subsystem> Subsystems { get; set; } = new List<Subsystem>();

            public SystemClass(string name)
            {
                Name = name ?? throw new ArgumentNullException(nameof(name));
            }

            public void AddSubsystem(Subsystem subsystem)
            {
                if (subsystem == null) throw new ArgumentNullException(nameof(subsystem));
                Subsystems.Add(subsystem);
            }

            public Subsystem GetOrCreateSubsystem(string subsystemName)
            {
                var subsys = Subsystems.FirstOrDefault(s => s.Name == subsystemName);
                if (subsys == null)
                {
                    subsys = new Subsystem(subsystemName);
                    Subsystems.Add(subsys);
                }
                return subsys;
            }

            public void MergeFrom(SystemClass other)
            {
                if (other == null) throw new ArgumentNullException(nameof(other));

                foreach (var otherSubsystem in other.Subsystems)
                {
                    // Находим или создаём подсистему
                    var existingSubsystem = Subsystems.FirstOrDefault(s => s.Name == otherSubsystem.Name);
                    if (existingSubsystem == null)
                    {
                        // Если такой подсистемы нет — добавляем всю
                        Subsystems.Add(otherSubsystem);
                        continue;
                    }

                    // Если подсистема есть — сливаем трубопроводные линии
                    foreach (var otherPipeLine in otherSubsystem.PipeLines)
                    {
                        var existingPipeLine = existingSubsystem.PipeLines
                            .FirstOrDefault(p => p.Name == otherPipeLine.Name);

                        if (existingPipeLine == null)
                        {
                            // Если линии нет — добавляем всю
                            existingSubsystem.PipeLines.Add(otherPipeLine);
                        }
                        else
                        {
                            // Если линия есть — просто добавляем её элементы
                            existingPipeLine.Elements.AddRange(otherPipeLine.Elements);
                        }
                    }
                }
            }

        }

    }

    public class ForColorModelElementId : IModelElementId
    {
        public Guid ElementId { get; }
        public Guid ModelPartId { get; }

        public ForColorModelElementId(Guid elementId, Guid modelPartId)
        {
            ElementId = elementId;
            ModelPartId = modelPartId;
        }
    }
}
