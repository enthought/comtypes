import abc
import io
import keyword
from typing import (
    Any,
    Dict,
    Generic,
    Iterable,
    Iterator,
    List,
    Optional,
    Sequence,
    Tuple,
    TYPE_CHECKING,
    TypeVar,
)

from comtypes.tools import typedesc

if TYPE_CHECKING:
    from comtypes import hints


if TYPE_CHECKING:
    _T_MTD = TypeVar("_T_MTD", bound=hints._MethodTypeDesc)
else:
    _T_MTD = TypeVar("_T_MTD")


class _MethodAnnotator(abc.ABC, Generic[_T_MTD]):
    def __init__(self, method: _T_MTD) -> None:
        self.method = method

    @property
    def inarg_specs(self) -> Sequence[Tuple[Any, str, Optional[Any]]]:
        index = 0
        result = []
        for typ, name, flags, default in self.method.arguments:
            if "in" in flags and "lcid" not in flags or not flags:
                index += 1
                if "optional" in flags:
                    default = ...
                result.append((typ, (name or f"__arg{index}"), default))
        return result

    @abc.abstractmethod
    def getvalue(self, name: str) -> str:
        ...


_CatMths = Tuple[  # categorized methods
    str, Optional[_T_MTD], Optional[_T_MTD], Optional[_T_MTD], Optional[_T_MTD]
]


class _MethodsAnnotator(abc.ABC, Generic[_T_MTD]):
    def __init__(self) -> None:
        self.stream = io.StringIO()

    @abc.abstractmethod
    def to_method_annotator(self, method: _T_MTD) -> _MethodAnnotator[_T_MTD]:
        ...

    def _iter_methods(self, members: Iterable[_T_MTD]) -> Iterator[_CatMths[_T_MTD]]:
        methods: Dict[str, List[Optional[_T_MTD]]] = {}
        MTH = 0
        GET = 1
        PUT = 2
        PUTREF = 3
        for mem in members:
            if "propget" in mem.idlflags:
                methods.setdefault(mem.name, [None] * 4)[GET] = mem
            elif "propput" in mem.idlflags:
                methods.setdefault(mem.name, [None] * 4)[PUT] = mem
            elif "propputref" in mem.idlflags:
                methods.setdefault(mem.name, [None] * 4)[PUTREF] = mem
            else:
                methods.setdefault(mem.name, [None] * 4)[MTH] = mem
        for name, (fmth, fget, fput, fputref) in methods.items():
            yield name, fmth, fget, fput, fputref

    def generate(self, members: Iterable[_T_MTD]) -> str:
        for name, fmth, fget, fput, fputref in self._iter_methods(members):
            if fmth:
                self._gen_method(name, fmth)
            elif fget and not fput and not fputref:
                self._gen_prop_get(name, fget)
            elif fget and fput and not fputref:
                self._gen_prop_get_put(name, fget, fput)
            elif fget and not fput and fputref:
                self._gen_prop_get_putref(name, fget, fputref)
            elif fget and fput and fputref:
                self._gen_prop_get_put_putref(name, fget, fput, fputref)
            elif not fget and fput and not fputref:
                self._gen_prop_put(name, fput)
            elif not fget and not fput and fputref:
                self._gen_prop_putref(name, fputref)
            elif not fget and fput and fputref:
                self._gen_prop_put_putref(name, fput, fputref)
            else:
                self._define_member(f"pass  # what does `{name}` behave?")
            self._patch_dunder(name)
        return self.stream.getvalue()

    def _patch_dunder(self, name: str) -> None:
        if name == "Count":
            self._define_member(f"__len__ = hints.to_dunder_len({name})")
        if name == "Item":
            self._define_member(f"__call__ = hints.to_dunder_call({name})")
            self._define_member(f"__getitem__ = hints.to_dunder_getitem({name})")
            self._define_member(f"__setitem__ = hints.to_dunder_setitem({name})")
        if name == "_NewEnum":
            self._define_member(f"__iter__ = hints.to_dunder_iter({name})")

    def _define_named_prop(
        self, mem_name: str, getter: Optional[str] = None, setter: Optional[str] = None
    ) -> None:
        if getter and setter:
            content = (
                f"{mem_name} = hints.named_property('{mem_name}', {getter}, {setter})"
            )
        elif getter and not setter:
            content = f"{mem_name} = hints.named_property('{mem_name}', {getter})"
        elif not getter and setter:
            content = f"{mem_name} = hints.named_property('{mem_name}', fset={setter})"
        else:
            return
        self._define_member(content)

    def _define_normal_prop(
        self, mem_name: str, getter: Optional[str] = None, setter: Optional[str] = None
    ) -> None:
        if getter and setter:
            content = f"{mem_name} = hints.normal_property({getter}, {setter})"
        elif getter and not setter:
            content = f"{mem_name} = hints.normal_property({getter})"
        elif not getter and setter:
            content = f"{mem_name} = hints.normal_property(fset={setter})"
        else:
            return
        self._define_member(content)

    def _define_member(self, content: str) -> None:
        print(f"        {content}", file=self.stream)

    def _gen_method(self, name: str, mth: _T_MTD) -> None:
        self._define_member(self.to_method_annotator(mth).getvalue(name))

    def _gen_prop_get(self, name: str, fget: _T_MTD) -> None:
        getter_anno = self.to_method_annotator(fget)
        self._define_member(getter_anno.getvalue(f"_get_{name}"))
        if getter_anno.inarg_specs:
            self._define_named_prop(name, f"_get_{name}")
        else:
            self._define_normal_prop(name, f"_get_{name}")

    def _gen_prop_get_put(self, name: str, fget: _T_MTD, fput: _T_MTD) -> None:
        getter_anno = self.to_method_annotator(fget)
        setter_anno = self.to_method_annotator(fput)
        self._define_member(getter_anno.getvalue(f"_get_{name}"))
        self._define_member(setter_anno.getvalue(f"_set_{name}"))
        if getter_anno.inarg_specs:
            self._define_named_prop(name, f"_get_{name}", f"_set_{name}")
        else:
            self._define_normal_prop(name, f"_get_{name}", f"_set_{name}")

    def _gen_prop_get_putref(self, name: str, fget: _T_MTD, fputref: _T_MTD) -> None:
        getter_anno = self.to_method_annotator(fget)
        setter_anno = self.to_method_annotator(fputref)
        self._define_member(getter_anno.getvalue(f"_get_{name}"))
        self._define_member(setter_anno.getvalue(f"_setref_{name}"))
        if getter_anno.inarg_specs:
            self._define_named_prop(name, f"_get_{name}", f"_setref_{name}")
        else:
            self._define_normal_prop(name, f"_get_{name}", f"_setref_{name}")

    def _gen_prop_get_put_putref(
        self, name: str, fget: _T_MTD, fput: _T_MTD, fputref: _T_MTD
    ) -> None:
        getter_anno = self.to_method_annotator(fget)
        put_anno = self.to_method_annotator(fput)
        putref_anno = self.to_method_annotator(fputref)
        self._define_member(getter_anno.getvalue(f"_get_{name}"))
        self._define_member(put_anno.getvalue(f"_set_{name}"))
        self._define_member(putref_anno.getvalue(f"_setref_{name}"))
        setter = f"hints.put_or_putref(_set_{name}, _setref_{name})"
        if getter_anno.inarg_specs:
            self._define_named_prop(name, f"_get_{name}", setter)
        else:
            self._define_normal_prop(name, f"_get_{name}", setter)

    def _gen_prop_put(self, name: str, fput: _T_MTD) -> None:
        setter_anno = self.to_method_annotator(fput)
        self._define_member(setter_anno.getvalue(f"_set_{name}"))
        if len(setter_anno.inarg_specs) >= 2:
            self._define_named_prop(name, setter=f"_set_{name}")
        else:
            self._define_normal_prop(name, setter=f"_set_{name}")

    def _gen_prop_putref(self, name: str, fputref: _T_MTD) -> None:
        setter_anno = self.to_method_annotator(fputref)
        self._define_member(setter_anno.getvalue(f"_setref_{name}"))
        if len(setter_anno.inarg_specs) >= 2:
            self._define_named_prop(name, setter=f"_setref_{name}")
        else:
            self._define_normal_prop(name, setter=f"_setref_{name}")

    def _gen_prop_put_putref(self, name: str, fput: _T_MTD, fputref: _T_MTD) -> None:
        put_anno = self.to_method_annotator(fput)
        putref_anno = self.to_method_annotator(fputref)
        self._define_member(put_anno.getvalue(f"_set_{name}"))
        self._define_member(putref_anno.getvalue(f"_setref_{name}"))
        setter = f"hints.put_or_putref(_set_{name}, _setref_{name})"
        if len(put_anno.inarg_specs) >= 2:
            self._define_named_prop(name, setter=setter)
        else:
            self._define_normal_prop(name, setter=setter)


class ComMethodAnnotator(_MethodAnnotator[typedesc.ComMethod]):
    def _iter_outarg_specs(self) -> Iterator[Tuple[Any, str]]:
        for typ, name, flags, _ in self.method.arguments:
            if "out" in flags:
                yield typ, name

    def getvalue(self, name: str) -> str:
        inargs = []
        has_optional = False
        for _, argname, default in self.inarg_specs:
            if keyword.iskeyword(argname):
                inargs = ["*args: Any", "**kwargs: Any"]
                break
            if default is None:
                if has_optional:
                    # probably propput or propputref
                    # HACK: Something that goes into this conditional branch
                    #       should be a special callback.
                    inargs.append("**kwargs: Any")
                    break
                inargs.append(f"{argname}: Any")
            else:
                inargs.append(f"{argname}: Any = ...")
                has_optional = True
        outargs = ["Any" for _ in self._iter_outarg_specs()]
        if not outargs:
            out = "Any"
        elif len(outargs) == 1:
            out = outargs[0]
        else:
            out = "Tuple[" + ", ".join(outargs) + "]"
        in_ = ("self, " + ", ".join(inargs)) if inargs else "self"
        return f"def {name}({in_}) -> {out}: ..."


class ComMethodsAnnotator(_MethodsAnnotator[typedesc.ComMethod]):
    def to_method_annotator(self, m: typedesc.ComMethod) -> ComMethodAnnotator:
        return ComMethodAnnotator(m)


class ComInterfaceMembersAnnotator(object):
    def __init__(self, itf: typedesc.ComInterface):
        self.itf = itf

    def generate(self) -> str:
        return ComMethodsAnnotator().generate(self.itf.members)


class DispMethodAnnotator(_MethodAnnotator[typedesc.DispMethod]):
    def getvalue(self, name: str) -> str:
        inargs = []
        has_optional = False
        for _, argname, default in self.inarg_specs:
            if keyword.iskeyword(argname):
                inargs = ["*args: Any", "**kwargs: Any"]
                break
            if default is None:
                if has_optional:
                    # probably propput or propputref
                    # HACK: Something that goes into this conditional branch
                    #       should be a special callback.
                    inargs.append("**kwargs: Any")
                    break
                inargs.append(f"{argname}: Any")
            else:
                inargs.append(f"{argname}: Any = ...")
                has_optional = True
        out = "Any"
        in_ = ("self, " + ", ".join(inargs)) if inargs else "self"
        return f"def {name}({in_}) -> {out}: ..."


class DispMethodsAnnotator(_MethodsAnnotator[typedesc.DispMethod]):
    def to_method_annotator(self, m: typedesc.DispMethod) -> DispMethodAnnotator:
        return DispMethodAnnotator(m)


class DispInterfaceMembersAnnotator(object):
    def __init__(self, itf: typedesc.DispInterface):
        self.itf = itf
        self.stream = io.StringIO()

    def _categorize_members(
        self,
    ) -> Tuple[Iterable[typedesc.DispProperty], Iterable[typedesc.DispMethod]]:
        props: List[typedesc.DispProperty] = []
        methods: List[typedesc.DispMethod] = []
        for mem in self.itf.members:
            if isinstance(mem, typedesc.DispMethod):
                methods.append(mem)
            elif isinstance(mem, typedesc.DispProperty):
                props.append(mem)
        return props, methods

    def generate(self) -> str:
        props, methods = self._categorize_members()
        for mem in props:
            print("        @property  # dispprop", file=self.stream)
            print(f"        def {mem.name}(self) -> Any: ...", file=self.stream)
        return self.stream.getvalue() + DispMethodsAnnotator().generate(methods)
