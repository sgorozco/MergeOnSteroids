using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Text.Json.Serialization;

namespace MergeOnSteroids.Core.Blocks;

/// <summary>
/// Base class for every "scratch piece". Blocks form a tree: container blocks
/// (loops, conditionals, documents) own one or more <see cref="BlockCollection"/>s.
/// </summary>
[JsonPolymorphic(TypeDiscriminatorPropertyName = "$type")]
[JsonDerivedType(typeof(CsvSourceBlock), "csvSource")]
[JsonDerivedType(typeof(ExcelSourceBlock), "excelSource")]
[JsonDerivedType(typeof(DatabaseSourceBlock), "databaseSource")]
[JsonDerivedType(typeof(FilterSourceBlock), "filterSource")]
[JsonDerivedType(typeof(ForEachBlock), "forEach")]
[JsonDerivedType(typeof(IfBlock), "if")]
[JsonDerivedType(typeof(SwitchBlock), "switch")]
[JsonDerivedType(typeof(CaseBlock), "case")]
[JsonDerivedType(typeof(SetVariableBlock), "setVariable")]
[JsonDerivedType(typeof(MakeDirectoryBlock), "makeDirectory")]
[JsonDerivedType(typeof(ZipDirectoryBlock), "zipDirectory")]
[JsonDerivedType(typeof(NewDocumentBlock), "newDocument")]
[JsonDerivedType(typeof(ParagraphBlock), "paragraph")]
[JsonDerivedType(typeof(WordFragmentBlock), "wordFragment")]
[JsonDerivedType(typeof(TableBlock), "table")]
[JsonDerivedType(typeof(PageBreakBlock), "pageBreak")]
public abstract class Block : INotifyPropertyChanged
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");

    /// <summary>The collection this block currently lives in (maintained by BlockCollection).</summary>
    [JsonIgnore]
    public BlockCollection? ParentCollection { get; internal set; }

    [JsonIgnore] public abstract string DisplayName { get; }

    public event PropertyChangedEventHandler? PropertyChanged;

    protected bool Set<T>(ref T field, T value, [CallerMemberName] string? name = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value)) return false;
        field = value;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        return true;
    }

    protected void Raise(string propertyName) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

    /// <summary>All child collections (containers override).</summary>
    public virtual IEnumerable<BlockCollection> ChildLists() => [];

    /// <summary>Depth-first enumeration of this block and all descendants.</summary>
    public IEnumerable<Block> SelfAndDescendants()
    {
        yield return this;
        foreach (var list in ChildLists())
            foreach (var child in list)
                foreach (var d in child.SelfAndDescendants())
                    yield return d;
    }

    /// <summary>True if <paramref name="other"/> is this block or nested anywhere inside it.</summary>
    public bool ContainsBlock(Block other)
    {
        var col = other.ParentCollection;
        while (col is not null)
        {
            if (ReferenceEquals(col.Owner, this)) return true;
            col = col.Owner?.ParentCollection;
        }
        return ReferenceEquals(this, other);
    }
}

/// <summary>
/// Observable list of blocks that keeps each child's <see cref="Block.ParentCollection"/> in sync,
/// so drag &amp; drop / delete can always find where a block lives.
/// </summary>
public class BlockCollection : ObservableCollection<Block>
{
    [JsonIgnore] public Block? Owner { get; internal set; }
    [JsonIgnore] public string SlotName { get; internal set; } = "";

    public BlockCollection() { }

    public BlockCollection(Block? owner, string slotName)
    {
        Owner = owner;
        SlotName = slotName;
    }

    protected override void InsertItem(int index, Block item)
    {
        base.InsertItem(index, item);
        item.ParentCollection = this;
    }

    protected override void SetItem(int index, Block item)
    {
        this[index].ParentCollection = null;
        base.SetItem(index, item);
        item.ParentCollection = this;
    }

    protected override void RemoveItem(int index)
    {
        this[index].ParentCollection = null;
        base.RemoveItem(index);
    }

    protected override void ClearItems()
    {
        foreach (var b in this) b.ParentCollection = null;
        base.ClearItems();
    }
}
