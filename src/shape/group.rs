//! `Group`：组合形状。
//!
//! 组合是把多个形状"打包"为一个整体——移动、旋转、缩放时内部子形状按
//! 相对位置一起变换。OOXML 用 `<p:grpSp>` 容器 + `<a:chOff>` / `<a:chExt>`
//! 表达子坐标系。
//!
//! # 与 python-pptx 的对应
//!
//! - `pptx.shapes.group.Group` ←→ [`Group`]；
//! - `pptx.shapes.group.GroupShape`（与 `Shape` 同源）←→ [`GroupChild`]。
//!
//! # 坐标系
//!
//! Group 自身有 `off` / `ext` 表示"组合在父 slide 中的位置 + 尺寸"；
//! `chOff` / `chExt` 是子坐标系（默认与 off/ext 相同）。
//! 修改 group 自身位置时同时 mutate `self.group.off`（参见 [`Shape::set_left`] 实现）。
//!
//! # 限制
//!
//! - 递归 Group（Group 嵌套 Group）已支持；
//! - 子形状可读取，也可通过 `Group::add_child` / `Group::remove_child` 增删。

use crate::oxml::shape::Group as OxmlGroup;
use crate::oxml::shape::GroupChild as OxmlGroupChild;
use crate::shape::autoshape::AutoShape;
use crate::shape::base::Shape;
use crate::shape::connector::Connector;
use crate::shape::picture::Picture;
use crate::shape::table::TableShape;
use crate::units::Emu;

/// 组合形状。
#[derive(Clone, Debug, Default)]
pub struct Group {
    /// 内部 oxml 句柄。
    pub(crate) group: OxmlGroup,
}

impl Group {
    /// 从 oxml 构造。
    pub fn from_group(g: OxmlGroup) -> Self {
        Group { group: g }
    }
    /// 取出 oxml 引用。
    pub fn group(&self) -> &OxmlGroup {
        &self.group
    }
    /// 取出 oxml 可变引用。
    pub fn group_mut(&mut self) -> &mut OxmlGroup {
        &mut self.group
    }

    /// 取所有子形状（递归）。
    pub fn children(&self) -> Vec<GroupChild> {
        self.group
            .children
            .iter()
            .map(|c| match c {
                OxmlGroupChild::Sp(s) => GroupChild::Sp(AutoShape::from_sp(s.clone())),
                OxmlGroupChild::Pic(p) => GroupChild::Pic(Picture::from_pic(p.clone())),
                OxmlGroupChild::CxnSp(c) => GroupChild::Cx(Connector::from_cxn(c.clone())),
                OxmlGroupChild::Group(g) => {
                    GroupChild::Grp(Box::new(Group::from_group((**g).clone())))
                }
                // 图形框：目前仅支持 Table 高阶句柄，其它类型（chart/ole/smartArt）走默认 TableShape 占位。
                // 读路径已通过 Graphic::SmartArt 保留原始 XML（TODO-037），但高阶层访问仍需通过 TableShape.frame.graphic。
                OxmlGroupChild::GraphicFrame(g) => {
                    GroupChild::Gfx(TableShape::from_frame(g.clone()))
                }
            })
            .collect()
    }

    /// 子形状数量。
    pub fn len(&self) -> usize {
        self.group.children.len()
    }
    /// 是否无子形状。
    pub fn is_empty(&self) -> bool {
        self.group.children.is_empty()
    }

    // --------------------- 子形状编辑 API（TODO-032 高阶） ---------------------
    //
    // 对标 python-pptx 中通过 `group.shapes._spTree.append/remove` 操作子形状。
    // 本库在 `Group` 上提供类型安全的 add/remove 接口，调用方传入高阶形状，
    // 自动转换为底层 OxmlGroupChild 后追加。

    /// 追加一个 [`AutoShape`] 子形状到组合末尾。
    pub fn add_autoshape(&mut self, shape: AutoShape) -> &mut Self {
        self.group
            .children
            .push(OxmlGroupChild::Sp(shape.sp.clone()));
        self
    }

    /// 追加一个 [`Picture`] 子形状到组合末尾。
    pub fn add_picture(&mut self, pic: Picture) -> &mut Self {
        self.group
            .children
            .push(OxmlGroupChild::Pic(pic.pic.clone()));
        self
    }

    /// 追加一个 [`Connector`] 子形状到组合末尾。
    pub fn add_connector(&mut self, cxn: Connector) -> &mut Self {
        self.group
            .children
            .push(OxmlGroupChild::CxnSp(cxn.cxn.clone()));
        self
    }

    /// 追加一个 [`TableShape`]（GraphicFrame）子形状到组合末尾。
    pub fn add_table(&mut self, table: TableShape) -> &mut Self {
        self.group
            .children
            .push(OxmlGroupChild::GraphicFrame(table.frame.clone()));
        self
    }

    /// 追加一个嵌套 [`Group`] 子形状到组合末尾。
    pub fn add_group(&mut self, grp: Group) -> &mut Self {
        self.group
            .children
            .push(OxmlGroupChild::Group(Box::new(grp.group.clone())));
        self
    }

    /// 按形状 ID 移除子形状。返回是否移除成功。
    ///
    /// 递归匹配：如果第一层未找到该 ID，会继续在嵌套 Group 中查找并移除。
    pub fn remove_child(&mut self, id: u32) -> bool {
        // 先在第一层找
        let pos = self.group.children.iter().position(|c| match c {
            OxmlGroupChild::Sp(s) => s.id == id,
            OxmlGroupChild::Pic(p) => p.id == id,
            OxmlGroupChild::CxnSp(c) => c.id == id,
            OxmlGroupChild::Group(g) => g.id == id,
            OxmlGroupChild::GraphicFrame(g) => g.id == id,
        });
        if let Some(p) = pos {
            self.group.children.remove(p);
            return true;
        }
        // 递归到嵌套 Group 中查找
        for c in &mut self.group.children {
            if let OxmlGroupChild::Group(g) = c {
                let mut sub = Group::from_group((**g).clone());
                if sub.remove_child(id) {
                    **g = sub.group;
                    return true;
                }
            }
        }
        false
    }

    /// 清空所有子形状。
    pub fn clear(&mut self) {
        self.group.children.clear();
    }
}

/// 高阶 GroupChild 枚举。
///
/// 与 [`crate::oxml::shape::GroupChild`] 的区别在于：本枚举承载的是
/// 高阶包装（`AutoShape` / `Picture` / `Connector` / `Group` / `TableShape`），
/// 方便调用方继续操作。
#[derive(Clone, Debug)]
pub enum GroupChild {
    /// 自选形状。
    Sp(AutoShape),
    /// 图片。
    Pic(Picture),
    /// 连接器。
    Cx(Connector),
    /// 递归 Group。
    Grp(Box<Group>),
    /// 图形框（目前仅承载表格）。
    Gfx(TableShape),
}

impl Shape for Group {
    fn id(&self) -> u32 {
        self.group.id
    }
    fn set_id(&mut self, id: u32) {
        self.group.id = id;
    }
    fn name(&self) -> &str {
        &self.group.name
    }
    fn set_name(&mut self, name: String) {
        self.group.name = name;
    }
    fn shape_type(&self) -> &'static str {
        "group"
    }

    fn left(&self) -> Emu {
        Emu::new(self.group.off.0.value())
    }
    fn set_left(&mut self, emu: Emu) {
        self.group.off.0 = emu;
    }
    fn top(&self) -> Emu {
        Emu::new(self.group.off.1.value())
    }
    fn set_top(&mut self, emu: Emu) {
        self.group.off.1 = emu;
    }
    fn width(&self) -> Emu {
        Emu::new(self.group.ext.0.value())
    }
    fn set_width(&mut self, emu: Emu) {
        self.group.ext.0 = emu;
    }
    fn height(&self) -> Emu {
        Emu::new(self.group.ext.1.value())
    }
    fn set_height(&mut self, emu: Emu) {
        self.group.ext.1 = emu;
    }

    fn rotation(&self) -> f64 {
        self.group.properties.rot_deg.unwrap_or(0.0)
    }
    fn set_rotation(&mut self, deg: f64) {
        self.group.properties.rot_deg = Some(deg);
        let rot = (deg * 60_000.0) as i32;
        self.group.properties.xfrm.rot = Some(rot);
    }
}
